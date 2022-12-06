import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import kotlin.math.abs

object AppExcel {

    @JvmStatic
    fun main(args: Array<String>) {
        val supplierFile = File("C:\\Users\\javaa\\Downloads\\supplier.xlsx")
        val receiverFile = File("C:\\Users\\javaa\\Downloads\\receiver-1.xlsx")
        val receiverOutFile = File("C:\\Users\\javaa\\Downloads\\receiver-out.xlsx")
        val supplierSumFile = File("C:\\Users\\javaa\\Downloads\\supplierSum.xlsx")

        val (headers, rows) = readFileAndReturnHeadersAndRows(supplierFile, 0)
        val (receiverHeaders, receiverRows) = readFileAndReturnHeadersAndRows(receiverFile, 2)

        createSumByMishloahExcel(headers, rows, supplierSumFile)

        compareFiles(receiverOutFile, receiverHeaders, receiverRows, headers, rows)
    }

    private fun compareFiles(
        receiverOutFile: File,
        receiverHeaders: List<String>,
        receiverRows: List<List<String?>>,
        headers: List<String>,
        rows: List<List<String?>>
    ) {
        val wbOut = XSSFWorkbook();
        val sheet = wbOut.createSheet("new sheet")
        sheet.ctWorksheet.sheetViews.getSheetViewArray(0).rightToLeft = true
        val fileOut = FileOutputStream(receiverOutFile)
        val hR = sheet.createRow(0)
        receiverHeaders.indices.forEach { i ->
            hR.createCell(i).setCellValue(receiverHeaders[i])
        }

        val mMismahIndex = takeIndex(receiverHeaders, "מ.מסמך")
        val dateIndex = takeIndex(receiverHeaders, "תאריך")
        val shemNekudaIndex = takeIndex(receiverHeaders, "שם הנקודה")
        val priceIndex = takeIndex(receiverHeaders, """לתשלום""")
        val sapakIndex = 6//takeIndex(receiverHeaders, """דרישת ספק""")
        //val mahMakorIndex = takeIndex(receiverHeaders, """מך מקור""")

        val supplierPointIndex = 9 //takeIndex(headers, )
        val supplierMishloahIndex = takeIndex(headers, "ת. משלוח")
        val supplierDateIndex = takeIndex(headers, "תאריך משלוח")
        val supplierPriceIndex = takeIndex(headers, """ס. עם מע"מ""")
        //val supplierAsmahtaIndex = takeIndex(headers, """ס. עם מע"מ""")

        var rowNum = 1
        receiverRows.forEach { r ->
            val currRow = sheet.createRow(rowNum++)
            r.indices.forEach { currRow.createCell(it).setCellValue(r[it]) }

            val price = r[priceIndex]
            if (!price.isNullOrBlank()) {
                val nearestAmount =
                    searchRowInSupplier(
                        r[mMismahIndex],
                        r[dateIndex],
                        r[shemNekudaIndex],
                        price,
                        rows,
                        supplierPointIndex,
                        supplierMishloahIndex,
                        supplierDateIndex,
                        supplierPriceIndex
                    )
                val drishatSapak = nearestAmount?.let { String.format("%.2f", it) }
                val leTashlum = price?.toDouble()
                currRow.createCell(sapakIndex).setCellValue(drishatSapak ?: "<NO>")
                if (nearestAmount != null && leTashlum != null) {
                    val diff = String.format("%.2f", leTashlum - nearestAmount)
                    currRow.createCell(sapakIndex + 1).setCellValue(diff)
                }
            }
        }


        wbOut.write(fileOut);
        fileOut.close()
    }

    private fun searchRowInSupplier(
        mishloah: String?,
        date: String?,
        nekuda: String?,
        price: String?,
        rows: List<List<String?>>,
        supplierPointIndex: Int,
        supplierMishloahIndex: Int,
        supplierDateIndex: Int,
        supplierPriceIndex: Int
    ): Double? {
        if (price.isNullOrEmpty()) return null
        val priceDouble = price.toDouble()
        val result = rows.filter { r ->
            val p = r[supplierPointIndex]
            val d = r[supplierDateIndex]

            if (priceDouble < 0) sameForNegative(nekuda, p) else same(date, nekuda, d, p)
        }
        val grouppedByMishloach = result.groupBy { it[supplierMishloahIndex] }
        val amountsGrouppedByMishloach =
            grouppedByMishloach.mapValues { (k, v) -> v.sumOf { it[supplierPriceIndex]?.toDouble() ?: 0.0 } }
        if (amountsGrouppedByMishloach.containsKey(mishloah)) {
            print("*")
            return amountsGrouppedByMishloach[mishloah]
        }
        print(grouppedByMishloach.size)
        val best = amountsGrouppedByMishloach.values.sortedBy { abs(priceDouble - it) }.firstOrNull()

        return if (comparePrices(best, price, 10)) best else null

    }

    private fun same(
        date: String?,
        nekuda: String?,
        d: String?,
        p: String?
    ): Boolean {
        val nekudaName = (nekuda?.split(" ") ?: listOf("<>"))[0]
        return date == d &&
                p?.contains(nekudaName) ?: false
    }

    private fun sameForNegative(
        nekuda: String?,
        p: String?
    ): Boolean {
        val nekudaName = (nekuda?.split(" ") ?: listOf("<>"))[0]
        return p?.contains(nekudaName) ?: false
    }

    private fun comparePrices(p1: Double?, p2: String?, percents: Int): Boolean {
        if (p1 == null || p2 == null)
            return false
        val diff = Math.abs(p1 - p2.toDouble()) / p1.toDouble()
        return diff <= percents / 100.0
    }

    private fun createSumByMishloahExcel(
        headers: List<String>,
        rows: List<List<String?>>,
        supplierSumFile: File
    ) {
        val requiredHeaders = listOf(
            "חשבונית",
            "ת. משלוח",
            "תיאור",
            "ברקוד",
            "תיאור",
            "תאריך משלוח",
            "כ יחידות",
            "מ ליחידה",
            "ערך ברוטו",
            "% הנחה",
            "הנחה",
            "עם פיקדון",
            """ס. עם מע"מ"""
        )
        val indexes = requiredHeaders.map { takeIndex(headers, it) }.toIntArray()
        indexes[2] = 9 //because we have many such values
        indexes[4] = takeLastIndex(headers, "תיאור") //because we have many such values
        val mishloahIndex = indexes[1]
        val sum1IndexToWrite = requiredHeaders.indexOf("כ יחידות")
        val sum2IndexToWrite = requiredHeaders.indexOf("עם פיקדון")
        val sum3IndexToWrite = requiredHeaders.indexOf("""ס. עם מע"מ""")
        val sum1Index = indexes[sum1IndexToWrite]
        val sum2Index = indexes[sum2IndexToWrite]
        val sum3Index = indexes[sum3IndexToWrite]


        val groupedRows = rows.groupBy { it[mishloahIndex] }

        val wbOut = XSSFWorkbook();
        val sheet = wbOut.createSheet("new sheet")
        sheet.ctWorksheet.sheetViews.getSheetViewArray(0).rightToLeft = true
        val fileOut = FileOutputStream(supplierSumFile)
        val hR = sheet.createRow(0)
        requiredHeaders.indices.forEach { i ->
            hR.createCell(i).setCellValue(requiredHeaders[i])
        }
        var rowNum = 1
        groupedRows.forEach { r ->
            val orig = r.value
            orig.forEach { origRow ->
                val currRow = sheet.createRow(rowNum++)
                origRow.forEach {
                    var j = 0
                    indexes.forEach { i ->
                        val s = origRow[i]
                        currRow.createCell(j++).setCellValue(s)
                    }
                }

            }
            val sumRow = sheet.createRow(rowNum++)
            sumRow.createCell(sum1IndexToWrite).setCellValue(orig.sumOf {
                val s = it[sum1Index]
                if (s.isNullOrEmpty()) {
                    println(orig)
                    0.0
                } else {
                    s.toDouble()
                }
            }
            )
            sumRow.createCell(sum2IndexToWrite).setCellValue(orig.sumOf {
                val s = it[sum2Index]
                if (s.isNullOrEmpty()) {
                    println(orig)
                    0.0
                } else {
                    s.toDouble()
                }
            })
            sumRow.createCell(sum3IndexToWrite).setCellValue(orig.sumOf {
                val s = it[sum3Index]
                if (s.isNullOrEmpty()) {
                    println(orig)
                    0.0
                } else {
                    s.toDouble()
                }
            })
        }
        wbOut.write(fileOut);
        fileOut.close()
    }

    private fun readFileAndReturnHeadersAndRows(
        supplierFile: File,
        headerLineNum: Int
    ): Pair<List<String>, List<List<String?>>> {
        val inputStream = FileInputStream(supplierFile)
        //Instantiate Excel workbook using existing file:
        val xlWb = WorkbookFactory.create(inputStream)

        //Row index specifies the row in the worksheet (starting at 0):

        //Get reference to first sheet:
        val xlWs = xlWb.getSheetAt(0)
        val firstRowNumber = xlWs.firstRowNum
        val lastRowNum = xlWs.lastRowNum
        val row = xlWs.getRow(headerLineNum)
        val firstColumn = row.firstCellNum
        val lastColumn = row.lastCellNum
        val headers =
            (firstRowNumber..lastRowNum + 1).map { row.getCell(it) }.filterNotNull().map { it.stringCellValue }


        val rows = (headerLineNum + 1..lastRowNum + 1).map { xlWs.getRow(it) }.filterNotNull()
            .filter { it.firstCellNum >= 0 && it.getCell(it.firstCellNum.toInt()).toString().isNotBlank() }
            .map { c ->
                (firstColumn..lastColumn + 1).map { c.getCell(it)?.toString() }

            }
        inputStream.close()
        return Pair(headers, rows)
    }


    private fun takeLastIndex(headers: List<String>, s: String): Int {
        return headers.lastIndexOf(s)
    }

    private fun takeIndex(headers: List<String>, s: String): Int {
        return headers.indexOf(s)
    }

    private fun toNum(s: String): Any {
        return s.toDoubleOrNull() ?: s
    }
}