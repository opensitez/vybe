// vybe-test: kotlin/data_classes/test_data_class_structural_equality_uses_all_fields
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Key(val id: Int, val tag: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Key(1, "x")
            val b = Key(1, "x")
            val c = Key(1, "y")
            __check((a == b).toString(), "true")
            __check((a == c).toString(), "false")
        }
