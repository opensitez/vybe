// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_aggregate_chars
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("apple", "ape", "bat", "ball", "cat")
            val out = words.groupingBy { it.first() }
                .aggregate { key, accumulator, element, first ->
                    val value = (accumulator ?: "") + element.length.toString()
                    value
                }
            __check((out['a']).toString(), "42")
            __check((out['b']).toString(), "12")
            __check((out['c']).toString(), "3")
        }
