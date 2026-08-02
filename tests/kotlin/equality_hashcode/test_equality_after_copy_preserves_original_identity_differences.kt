// vybe-test: kotlin/equality_hashcode/test_equality_after_copy_preserves_original_identity_differences
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Point(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Point(1, 2)
            val copy = base.copy()
            __check((base == copy).toString(), "true")
            __check((base === copy).toString(), "false")
            __check((base.x == copy.x).toString(), "true")
        }
