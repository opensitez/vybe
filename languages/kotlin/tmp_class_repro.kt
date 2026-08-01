class Logger {
    private var tag = "log"
    inner class Entry {
        fun marker(): Logger = this@Logger
    }

    fun tag(): String = tag
}

fun main() {
    val entry = Logger().Entry()
    println(entry.marker().tag())
}
