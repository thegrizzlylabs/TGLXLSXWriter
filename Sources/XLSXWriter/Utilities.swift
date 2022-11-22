import Foundation

/// Inspired by https://github.com/Kitura/swift-html-entities
extension String {
    func xmlEscape() -> String {
        // result buffer
        var str: String = ""

        for c in self {
            let unicodes = String(c).unicodeScalars

            for scalar in unicodes {
                let unicode = scalar.value

                if unicode.isSafeASCII {
                    str += String(scalar)
                } else {
                    let codeStr = String(unicode, radix: 10)

                    str += "&#" + codeStr + ";"
                }
            }
        }

        return str
    }
}

extension UInt32 {
    var isASCII: Bool {
        // Less than 0x80
        return self < 0x80
    }

    /// https://www.w3.org/International/questions/qa-escapes#use
    var isAttributeSyntax: Bool {
        // unicode values of [", ']
        return self == 0x22 || self == 0x27
    }

    /// https://www.w3.org/International/questions/qa-escapes#use
    var isTagSyntax: Bool {
        // unicode values of [&, < , >]
        return self.isAmpersand || self == 0x3C || self == 0x3E
    }

    var isAmpersand: Bool {
        // unicode value of &
        return self == 0x26
    }

    var isSafeASCII: Bool {
        return self.isASCII && !self.isAttributeSyntax && !self.isTagSyntax
    }
}
