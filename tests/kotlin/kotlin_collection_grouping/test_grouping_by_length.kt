// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_length
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("a", "bee", "cat", "deer", "dog")
            val grouped = words.groupBy { it.length }
            val short = grouped[1]!!.joinToString(",")
            val long = grouped[3]!!.joinToString(",")
            __check((short).toString(), "a")
            __check((long).toString(), "cat,dog")
        }
