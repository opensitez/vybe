// vybe-test: kotlin/equality_hashcode/test_mutable_list_component_equality_behavior
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf(1, 2)
            val right = listOf(1, 2)
            __check((left == right).toString(), "true")
            __check((left === right).toString(), "false")
            __check((left.hashCode() == right.hashCode()).toString(), "true")
        }
