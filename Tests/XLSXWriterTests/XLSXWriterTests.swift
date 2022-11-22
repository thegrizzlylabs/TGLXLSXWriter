import XCTest
@testable import XLSXWriter

final class XLSXWriterTests: XCTestCase {
    func testBasic() throws {
        let writer = XLSXWriter(application: "tests")
        let tempURL = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent(UUID().uuidString).appendingPathExtension("xlsx")
        print(tempURL)
        writer.addRow(["Date", "Merchant", "Amount"])
        writer.addRow(["2022-01-01", "Auchan", "2 €"])
        writer.addRow(["2022-01-02", "Craco", "2", "3"])
        try writer.generate(outputFileURL: tempURL)
    }

    func testCellName() {
        let writer = XLSXWriter(application: "tests")
        XCTAssertEqual(writer.cellName(row: 0, col: 0), "A1")
        XCTAssertEqual(writer.cellName(row: 0, col: 1), "B1")
        XCTAssertEqual(writer.cellName(row: 0, col: 25), "Z1")
        XCTAssertEqual(writer.cellName(row: 0, col: 26), "AA1")
        XCTAssertEqual(writer.cellName(row: 1, col: 27), "AB2")
    }
}
