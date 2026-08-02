// vybe-test: kotlin/kotlin_set_apis/test_mutable_set_add_duplicate
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = mutableSetOf("x")
            val first = set.add("x")
            val second = set.add("y")
            __check((first).toString(), "false")
            __check((second).toString(), "true")
            __check((set.size).toString(), "2")
        }
