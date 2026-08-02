// vybe-test: kotlin/advanced_features/test_advanced_data_class_destructure_and_mutate_copy
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

data class Pair(val left: Int, val right: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val original = Pair(10, 20)
            val (a, b) = original
            val updated = original.copy(right = 99)
            __check((a).toString(), "10")
            __check((b).toString(), "20")
            __check((updated.left).toString(), "10")
            __check((updated.right).toString(), "99")
        }
