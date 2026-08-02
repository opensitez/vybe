// vybe-test: kotlin/data_classes/test_data_class_local_scope_declaration
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

fun make(): Int {
            data class Local(val value: Int)
            return Local(9).value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((make()).toString(), "9")
        }
