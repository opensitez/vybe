// vybe-test: kotlin/annotations/test_annotation_on_interface_member
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

interface Logger { @Deprecated("old") fun emit(): Int }
        class Console : Logger { override fun emit(): Int = 3 }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val l: Logger = Console()
__check((l.emit()).toString(), "3") }
