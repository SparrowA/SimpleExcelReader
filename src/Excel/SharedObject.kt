import java.util.*

/**
 * Object with high-function
 */
object SharedObject {
    fun Scanner.nextLine(skip : Int) : String {
        var skipped = 0
        var result = ""
        while(hasNextLine() && skipped <= skip) {
            result = nextLine()
            skipped++
        }
        return result
    }
}