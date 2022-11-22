import Foundation

public struct Row: ExpressibleByArrayLiteral {
    public typealias ArrayLiteralElement = String

    let cells: [String]

    public init(arrayLiteral elements: ArrayLiteralElement...) {
        self.cells = elements
    }
}
