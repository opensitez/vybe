// vybe-test: kotlin/kotlin_set_apis/test_set_partition_like
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(1, 2, 3, 4)
            val out = set.partition { it % 2 == 0 }
            __check((out.first.joinToString(",")).toString(), "2,4")
            __check((out.second.joinToString(",")).toString(), "1,3")
        }
