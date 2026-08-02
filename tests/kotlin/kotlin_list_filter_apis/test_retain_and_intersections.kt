// vybe-test: kotlin/kotlin_list_filter_apis/test_retain_and_intersections
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf(1, 2, 3, 4)
            val evens = source.filter { it % 2 == 0 }
            val odds = source.filter { it % 2 == 1 }
            __check((evens.joinToString(",")).toString(), "2,4")
            __check((odds.joinToString(",")).toString(), "1,3")
            __check((source.intersect(evens.toSet()).joinToString(",")).toString(), "2,4")
        }
