// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_to_existing_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf("k", "kk", "kotlin")
            val out = linkedMapOf<Int, MutableList<String>>()
            source.groupByTo(out, { it.length }, { it })
            __check((out[1]!!.joinToString(",")).toString(), "k")
            __check((out[6]!!.first()).toString(), "kotlin")
            __check((out.size).toString(), "3")
        }
