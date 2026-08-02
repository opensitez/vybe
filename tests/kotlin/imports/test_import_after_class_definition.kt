// vybe-test: kotlin/imports/test_import_after_class_definition
// origin: languages/kotlin/tests/kotlin/test_imports.rs

class Holder
        import kotlin.math.absoluteValue
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((-9).absoluteValue).toString(), "9")
        }
