// vybe-test: kotlin/equality_hashcode/test_classic_primitive_wrappers_stay_by_value
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left: Int? = 3
            val right: Int? = 3
            __check((left == right).toString(), "true")
            __check((left === right).toString(), "true")
        }
