import java.io.File

object App {
    @JvmStatic
    fun main(args: Array<String>) {
        val supplierFile = File("C:\\Users\\javaa\\Downloads\\supplier.csv")
        val receiverFile = File("C:\\Users\\javaa\\Downloads\\receiver.csv")

        val supplierLines = supplierFile.bufferedReader().readLines();
        val supplierHeaders = supplierLines[0].split(",")
        println(supplierHeaders)
        println(supplierLines[1].split(","))
    }
}