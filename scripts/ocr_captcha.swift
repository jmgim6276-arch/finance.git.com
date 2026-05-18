import AppKit
import CoreGraphics
import Foundation
import Vision

let expectedLength = 5
let scales = [4, 6, 8, 10, 12]
let thresholds = [210, 220, 230, 235, 240, 245, 250]

func makeCGImage(from image: NSImage) -> CGImage? {
    guard let tiff = image.tiffRepresentation,
          let bitmap = NSBitmapImageRep(data: tiff),
          let cgImage = bitmap.cgImage else {
        return nil
    }
    return cgImage
}

func normalizeCandidate(_ text: String) -> String {
    let filtered = text.lowercased().unicodeScalars.filter { CharacterSet.alphanumerics.contains($0) }
    return String(String.UnicodeScalarView(filtered))
}

func thresholdImage(_ image: CGImage, scale: Int, threshold: Int) -> CGImage? {
    let width = image.width * scale
    let height = image.height * scale
    let colorSpace = CGColorSpaceCreateDeviceRGB()
    guard let context = CGContext(
        data: nil,
        width: width,
        height: height,
        bitsPerComponent: 8,
        bytesPerRow: width * 4,
        space: colorSpace,
        bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
    ) else {
        return nil
    }

    context.setFillColor(NSColor.white.cgColor)
    context.fill(CGRect(x: 0, y: 0, width: width, height: height))
    context.interpolationQuality = .high
    context.draw(image, in: CGRect(x: 0, y: 0, width: width, height: height))

    guard let scaled = context.makeImage(),
          let data = scaled.dataProvider?.data,
          let ptr = CFDataGetBytePtr(data) else {
        return nil
    }

    let length = CFDataGetLength(data)
    let outData = UnsafeMutablePointer<UInt8>.allocate(capacity: length)
    for index in stride(from: 0, to: length, by: 4) {
        let red = Int(ptr[index])
        let green = Int(ptr[index + 1])
        let blue = Int(ptr[index + 2])
        let alpha = ptr[index + 3]
        let isInk = min(red, green, blue) < threshold
        let value: UInt8 = isInk ? 0 : 255
        outData[index] = value
        outData[index + 1] = value
        outData[index + 2] = value
        outData[index + 3] = alpha
    }

    guard let provider = CGDataProvider(
        dataInfo: nil,
        data: outData,
        size: length,
        releaseData: { _, data, _ in
            data.deallocate()
        }
    ) else {
        outData.deallocate()
        return scaled
    }

    return CGImage(
        width: width,
        height: height,
        bitsPerComponent: 8,
        bitsPerPixel: 32,
        bytesPerRow: width * 4,
        space: colorSpace,
        bitmapInfo: CGBitmapInfo(rawValue: CGImageAlphaInfo.premultipliedLast.rawValue),
        provider: provider,
        decode: nil,
        shouldInterpolate: false,
        intent: .defaultIntent
    )
}

func recognize(_ image: CGImage) -> String {
    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.usesLanguageCorrection = false
    request.minimumTextHeight = 0.05

    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    try? handler.perform([request])

    let joined = (request.results ?? [])
        .compactMap { $0.topCandidates(1).first?.string }
        .joined(separator: "")
    return normalizeCandidate(joined)
}

func emit(_ object: [String: Any]) {
    guard let data = try? JSONSerialization.data(withJSONObject: object, options: []),
          let text = String(data: data, encoding: .utf8) else {
        fputs("{\"ok\":false,\"error\":\"无法序列化 OCR 结果\"}\n", stderr)
        exit(2)
    }
    print(text)
}

guard CommandLine.arguments.count >= 2 else {
    emit([
        "ok": false,
        "error": "missing-image-path",
    ])
    exit(1)
}

let imagePath = CommandLine.arguments[1]
guard let image = NSImage(contentsOfFile: imagePath),
      let cgImage = makeCGImage(from: image) else {
    emit([
        "ok": false,
        "error": "cannot-load-image",
    ])
    exit(1)
}

var votes: [String: Int] = [:]
for scale in scales {
    for threshold in thresholds {
        guard let processed = thresholdImage(cgImage, scale: scale, threshold: threshold) else {
            continue
        }
        let candidate = recognize(processed)
        if candidate.isEmpty {
            continue
        }
        votes[candidate, default: 0] += 1
    }
}

let sortedCandidates = votes
    .map { (value: $0.key, count: $0.value) }
    .sorted {
        let lhsExact = $0.value.count == expectedLength
        let rhsExact = $1.value.count == expectedLength
        if lhsExact != rhsExact {
            return lhsExact && !rhsExact
        }
        if $0.count != $1.count {
            return $0.count > $1.count
        }
        if $0.value.count != $1.value.count {
            return abs($0.value.count - expectedLength) < abs($1.value.count - expectedLength)
        }
        return $0.value < $1.value
    }

guard let best = sortedCandidates.first else {
    emit([
        "ok": false,
        "error": "no-candidate",
    ])
    exit(1)
}

emit([
    "ok": true,
    "expectedLength": expectedLength,
    "code": best.value,
    "candidates": sortedCandidates.map { ["value": $0.value, "count": $0.count] },
])
