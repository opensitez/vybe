// vybe-test: kotlin/kotlin_set_apis/test_set_linked_insertion_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = linkedSetOf("z", "a", "m", "a")
            __check((set.joinToString(",")).toString(), "z,a,m")
            __check((set.size).toString(), "3")
        }
