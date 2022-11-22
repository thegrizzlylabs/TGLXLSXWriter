import Foundation

public enum XLSXWriterError: Error {
    case ioError
    case compressionError(Error)
}
