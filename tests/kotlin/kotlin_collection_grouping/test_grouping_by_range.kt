// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1..10).toList()
            val grouped = values.groupBy { if (it <= 3) "small" else "big" }
            __check((grouped["small"]!!.sum()).toString(), "6")
            __check((grouped["big"]!!.size).toString(), "7")
        }
