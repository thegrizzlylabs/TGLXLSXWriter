import Foundation
import XCTest
@testable import XLSXWriter

final class UtilitiesTests: XCTestCase {
    func testXMLEscape() {
        XCTAssertEqual("".xmlEscape(), "")
        XCTAssertEqual("The Grizzly Labs".xmlEscape(), "The Grizzly Labs")
        XCTAssertEqual("13 € < 15 €".xmlEscape(), "13 &#8364; &#60; 15 &#8364;")
    }
}
