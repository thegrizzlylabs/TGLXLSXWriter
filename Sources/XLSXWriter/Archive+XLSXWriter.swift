import Foundation
import ZIPFoundation

extension Archive {
    func addDirectory(_ name: String) throws {
        try addEntry(with: name, type: .directory, uncompressedSize: Int64(0), provider: { (position: Int64, size) in
            return Data()
        })
    }

    func addEntry(_ name: String, string: String) throws {
        let data = Data(string.utf8)
        try addEntry(with: name, type: .file, uncompressedSize: Int64(data.count), provider: { (position: Int64, size) in
            let position = Int(position)
            return data.subdata(in: position..<position+size)
        })
    }
}
