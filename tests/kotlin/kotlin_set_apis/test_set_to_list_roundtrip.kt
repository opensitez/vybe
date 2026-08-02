// vybe-test: kotlin/kotlin_set_apis/test_set_to_list_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = linkedSetOf("a", "b", "c")
            val list = set.toList()
            __check((list.joinToString(",")).toString(), "a,b,c")
            __check((set.joinToString(",")).toString(), "a,b,c")
        }
