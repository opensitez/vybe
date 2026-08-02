// vybe-test: kotlin/equality_hashcode/test_data_class_copy_changes_only_targeted_fields
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Pair(val left: Int, val right: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = Pair(1, "a")
            val updated = source.copy(left = 3)
            __check((source.left).toString(), "1")
            __check((updated.left).toString(), "3")
            __check((updated.right).toString(), "a")
            __check((source.right).toString(), "a")
        }
