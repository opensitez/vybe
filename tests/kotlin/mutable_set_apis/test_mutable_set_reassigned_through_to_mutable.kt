// vybe-test: kotlin/mutable_set_apis/test_mutable_set_reassigned_through_to_mutable
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            val copy = values.toMutableSet()
            copy.add(3)
            __check((values.joinToString(",")).toString(), "1,2")
            __check((copy.joinToString(",")).toString(), "1,2,3")
        }
