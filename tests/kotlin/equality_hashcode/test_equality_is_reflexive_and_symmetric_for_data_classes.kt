// vybe-test: kotlin/equality_hashcode/test_equality_is_reflexive_and_symmetric_for_data_classes
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Point(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Point(1, 2)
            val second = Point(1, 2)
            __check((first == first).toString(), "true")
            __check((first == second).toString(), "true")
            __check((second == first).toString(), "true")
        }
