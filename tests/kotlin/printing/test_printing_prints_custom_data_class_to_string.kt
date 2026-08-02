// vybe-test: kotlin/printing/test_printing_prints_custom_data_class_to_string
// origin: languages/kotlin/tests/kotlin/test_printing.rs

data class Box(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box(42)).toString(), "Box(value=42)")
        }
