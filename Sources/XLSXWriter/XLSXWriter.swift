import Foundation
import HTMLSpecialCharacters
import ZIPFoundation

public class XLSXWriter {
    private let application: String
    private let creationDate: Date

    private var rows: [Row] = []

    public init(
        creationDate: Date = Date(),
        application: String
    ) {
        self.creationDate = creationDate
        self.application = application
    }

    struct SharedStringValue {
        var occurrence: Int
        var position: Int
    }

    private var sharedStrings: [String: SharedStringValue] = [:]

    public func generate(outputFileURL: URL) throws {
        guard let archive = Archive(url: outputFileURL, accessMode: .create) else {
            throw XLSXWriterError.ioError
        }

        do {
            //archive.addDirectory("docProps")
            try archive.addEntry("docProps/app.xml", string: appXML())
            try archive.addEntry("docProps/core.xml", string: coreXML())

            //archive.addDirectory("_rels")
            try archive.addEntry("_rels/.rels", string: relsXML())

            // archive.addDirectory("xl/worksheets")
            try archive.addEntry("xl/worksheets/sheet1.xml", string: sheetXML())
            try archive.addEntry("xl/workbook.xml", string: workbookXML())
            try archive.addEntry("xl/sharedStrings.xml", string: sharedStringsXML())

            try archive.addEntry("xl/_rels/workbook.xml.rels", string: workbookRelsXML())

            try archive.addEntry("[Content_Types].xml", string: contentTypesXML())
        } catch {
            throw XLSXWriterError.compressionError(error)
        }
    }

    public func addRow(_ row: Row) {
        rows.append(row)
    }

    private func sharedStringNo(_ string: String) -> Int {
        var sharedStringValue: SharedStringValue
        if let existingSharedStringValue = sharedStrings[string] {
            sharedStringValue = existingSharedStringValue
        } else {
            sharedStringValue = SharedStringValue(occurrence: 0, position: (sharedStrings.values.map(\.position).max() ?? -1) + 1)
        }

        sharedStringValue.occurrence += 1

        sharedStrings[string] = sharedStringValue

        return sharedStringValue.position
    }

    func cellName(row: Int, col: Int) -> String {
        var n = col

        var name = ""

        while n >= 0 {
            name = "\(UnicodeScalar(n % 26 + 0x41)!)\(name)"
            n = Int(n/26) - 1
        }

        return "\(name)\(row + 1)"
    }

    public func sheetXML() -> String {
        var rowsXML: [String] = []

        for (rowNo, row) in rows.enumerated() {
            var cellsXML: [String] = []

            for (colNo, cellValue) in row.cells.enumerated() {
                let fieldType = cellValue.isNumber ? "n" : "s";

                let fieldValueNo = sharedStringNo(cellValue);

                cellsXML.append(String(
                    format: ###"<c r="%@" t="%@"><v>%d</v></c>"###,
                    cellName(row: rowNo, col: colNo).escapeHTML,
                    fieldType,
                    fieldValueNo
                ))
            }

            rowsXML.append(String(
                format:
                    """
                    <row r="%@">
                    %@
                    </row>
                    """,
                String(rowNo + 1).escapeHTML,
                cellsXML.joined(separator: "\n")
            ))
        }

        return String(
            format:
                """
                <?xml version="1.0" encoding="utf-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheetData>
                  %@
                  </sheetData>
                </worksheet>
                """,
            rowsXML.joined(separator: "\n")
        );
    }

    public func sharedStringsXML() -> String {
        var sharedStringsXML: [String] = []

        for sharedString in sharedStrings.sorted(by: { $0.value.position < $1.value.position }).map(\.key) {
            sharedStringsXML.append(String(
                format: #"<si><t>%@</t></si>"#, sharedString.escapeHTML
            ))
        }

        return String(format:
                      """
                      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                      <sst count="%d" uniqueCount="%d" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      %@
                      </sst>
                      """,
                      sharedStrings.values.map(\.occurrence).reduce(0, +),
                      sharedStrings.count,
                      sharedStringsXML.joined(separator: "\n")
        )
    }


    public func workbookXML() -> String {
        return String(format:
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets>
              <sheet name="Sheet1" sheetId="1" r:id="rId1" />
              </sheets>
            </workbook>
            """
        )
    }

    public func contentTypesXML() -> String {
        return String(format:
            """
            <?xml version="1.0" encoding="UTF-8"?>
              <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
              <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
              <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
              <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
              </Types>
            """
        )
    }

    public func workbookRelsXML() -> String {
        return String(format:
                      """
                      <?xml version="1.0" encoding="UTF-8"?>
                      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                      </Relationships>
                      """
        )
    }

    public func appXML() -> String {
        return String(format:
              """
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
              <Application>\(application)</Application>
              </Properties>
              """
        )
    }

    public func coreXML() -> String {
        return String(format:
                      """
                      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                      <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                      <dcterms:created xsi:type="dcterms:W3CDTF">%@</dcterms:created>
                      <dc:creator>\(application)</dc:creator>
                      </cp:coreProperties>
                      """,
                      ISO8601DateFormatter().string(from: creationDate)
        )
    }

    public func relsXML() -> String {
        return String(format:
            """
            <?xml version="1.0" encoding="UTF-8"?>
              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
              <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
              </Relationships>
            """
        );
    }
}
