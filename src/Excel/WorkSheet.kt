import java.io.InputStream
import java.util.*
import Excel.SharedObject.nextLine

/**
 * Created by Gusev-Ab on 05.10.2018.
 */
class WorkSheet(sheetStream: InputStream, workbook: WorkBook) {

    class WorkSheetRowIterator(private val workSheet: WorkSheet, private var curRow: Int, private var curCol: Int) : Iterator<String> {

        override fun hasNext(): Boolean {
            return workSheet.getValue(curRow, curCol) != ""
        }

        override fun next(): String {
            var result = ""
            if(hasNext()) {
                result = workSheet.getValue(curRow, curCol)
                curRow++
            }

            return result
        }

        fun getColumn(colNum : Int) = workSheet.getValue(curRow, colNum)

        fun nextColumn() = {
            curRow++
            workSheet.getValue(curRow, curCol)
        }
    }

    private val charOffset = 64

    private val workBook : WorkBook = workbook

    /**
     * Cache contains data from sheet.
     * Each element in list represent row on sheet
     * Each element contains Map represent columns from row.
     */
    private val sheetData : List<Map<String, String>>

    init {
        sheetData = fetchSheetValue((Scanner(sheetStream)).nextLine(1))
    }

    /**
     * Fetch data from sheet
     * @param dataString XML string contains sheet's data
     * @return Return List contains Map which in Key is letter designation of column
     */
    private fun fetchSheetValue(dataString : String) =
            dataString.substringAfter("<sheetData>")
                    .substringBefore("</sheetData>")
                    .split(Regex("<row.*?>"))
                    .filter { x -> !x.isEmpty() }
                    .map { x -> x.replace(Regex("(<row.*?>)|(</row>)"), "") }
                    .map { x -> extractRowColumnMap(x) }


    /**
     * Extract value from columns of row
     * @param row String contains XML represent of row
     * @return Map in which key is letter designation of column, value is value of cell
     */
    private fun extractRowColumnMap(row : String) : Map<String, String> =
            row.split("<c")
                    .filter { x -> !x.isEmpty() }
                    .associate { x -> Pair(extractCellKey(x), extractCellValue(x)) }

    /**
     * Extract from XML represents cell key of cell
     * @param cellString XML represent of cell
     */
    private fun extractCellKey(cellString : String) = (Regex("r=\"[A-Z]+[0-9]+\"")).find(cellString).let { it?.value ?: "" }.replace(Regex("\"|r="), "")


    /**
     * Extract from XML represents cell value of cell
     */
    private fun extractCellValue(cellString : String) =
            Regex("<v>[-A-Za-z0-9_#]+</v>")
                    .find(cellString)
                    .let { it?.value ?: "" }
                    .replace(Regex("<v>|</v>"), "")
                    .let { if(isSharedValue(cellString)) workBook.getSharedValue(it.toInt()) else it }

    /**
     * Determine whether value is shared or not
     * @param cellString XML repesent cell
     * @return True if cell is shared, else return false
     */
    private fun isSharedValue(cellString : String) = Regex("t=\"s\"").containsMatchIn(cellString)


    /**
     * Return cell value from sheet using [sheetData] cache
     * @param row Number of row of cell
     * @param column Number of column of cell
     * @return If there is value in given [row] [column] pair then return that else return empty string
     */
    fun getValue(row : Int, column : Int) =
            if(sheetData.size > row) sheetData[row].let { it["${(column + charOffset).toChar()}${row + 1}"] ?: "" }
            else ""

    /**
     * Create iterator by Sheet
     * @param row Start number of row, default 1
     * @param column Start number of column default 1
     */
    fun getSheetRowIterator(row : Int = 1, column : Int = 1) = WorkSheetRowIterator(this, row, column)

}