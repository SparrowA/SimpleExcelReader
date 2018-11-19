import java.io.File
import java.io.FileNotFoundException
import java.io.InputStream
import java.util.*
import java.util.zip.ZipFile
import Excel.SharedObject.nextLine

/**
 * Created by Gusev-Ab on 05.10.2018.
 */
class WorkBook(private val excelPath: String) {

    class FileExtensionException(p0: String?) : RuntimeException(p0)
    class WorkSheetNotFound(p0: String?) : RuntimeException(p0)
    class NotCorrectSharedString(p0: String?) : RuntimeException(p0)

    /**
     * List of sheet from workbook.
     * Repressent like Map key is name of sheet and value is id of sheet
     */
    private val sheetList : HashMap<String, String>
    private val sharedValue : ArrayList<String>

    /**
     * Determine whether book is open
     */
    var isOpen : Boolean
       private set


    init {
        this.excelPath = excelPath
        checkFile()

        sheetList = HashMap()
        sharedValue = ArrayList()
        isOpen = false
    }

    /**
     * @param excelPath Path to excel
     * @param openWorkBook Need open workbook instantly, default true
     */
    constructor(excelPath : String, openWorkBook : Boolean = true) : this(excelPath) {
        if(openWorkBook) OpenWorkBook()
    }

    /**
     * Fetching list of sheet and shared String from Excel
     */
    fun OpenWorkBook() {
        if(!isOpen) {
            fetchSheetList()
            fetchSharedValue()
            isOpen = true
        }
    }

    /**
     * Clear list of sheet and all load shared string
     */
    fun CloseWorkBook() {
        if(isOpen) {
            sheetList.clear()
            sharedValue.clear()
        }
    }


    /**
     * Check existing file and it's extension
     */
    private fun checkFile() {
        val file = File(excelPath)
        if(!file.exists()) {
            throw FileNotFoundException("File does't found: $excelPath")
        }

        if(file.extension != "xlsx") {
            throw FileExtensionException("File has incorrect extension ${file.extension} instead of xlsx")
        }
    }

    /**
     * Create stream for reading data from object in Excel
     * @param entryName Name of object
     */
    fun getEntryStream(entryName : String) : InputStream {
        val zip = ZipFile(excelPath)
        return zip.getInputStream(zip.getEntry(entryName))
    }

    /**
     * Fetch list of sheet in Excel
     */
    private fun fetchSheetList() {
        val scanner = Scanner(getEntryStream("xl/workbook.xml"))
        val start = scanner.nextLine(1)

        sheetList.putAll(
                (start.split("<sheet").filter { x -> x.startsWith(" name") }.filter { x -> !x.contains("state=\"hidden\"") } as ArrayList)
                        .apply { replaceAll { x ->
                            x.replace(" name=\"", "")
                                    .replace(Regex("\" sheetId=\"\\d+\" r:id=\"rId"), "|")
                                    .replace("\"/>", "")
                                    .replace(Regex("</sheets>.*"), "")
                        } }.associateBy ({ x -> x.split("|")[0] } , { x -> x.split("|")[1]})
        )

        scanner.close()
    }

    /**
     * Fetch shared string from Excel
     */
    private fun fetchSharedValue() {
        val scanner = Scanner(getEntryStream("xl/sharedStrings.xml"))
        var start = scanner.nextLine(1)

        if(start != "") {
            do {
                sharedValue.addAll(
                        (start.split("<si>").filter { x -> x.endsWith("/si>") } as ArrayList).apply {replaceAll { x -> x.replace("</t></si>", "").replace(Regex("<t.*>"), "")}}
                )

                if(scanner.hasNextLine()) {
                    start = start.substringAfterLast("</si>").replace("</si><si>", "") + scanner.nextLine()
                }
                else {
                    start = start.substringAfterLast("</si></sst>")

                    if(start != "") {
                        throw NotCorrectSharedString("Shared string xml file is not valid")
                    }
                    break
                }
            }
            while (true)
        }

        scanner.close()
    }

    /**
     * Create WorkSheet instanse
     * @param sheetName Name of sheet
     * @return Instanse of [WorkSheet]
     * @throws WorkSheetNotFound if Excel does not contain given sheet
     */
    fun getWorkSheet(sheetName : String) : WorkSheet {
        if(sheetList.containsKey(sheetName)) {
            return WorkSheet(getEntryStream("xl/worksheets/sheet${sheetList[sheetName]}.xml"), this)
        }
        else {
            throw WorkSheetNotFound("Worksheet $sheetName has not found")
        }
    }

    /**
     * Give shared value by index
     * @param valueIndex of shared string
     * @return If there is shared string by given index then return it else return empty value
     */
    fun getSharedValue(valueIndex : Int) = if(sharedValue.size > valueIndex) sharedValue[valueIndex] else ""
}