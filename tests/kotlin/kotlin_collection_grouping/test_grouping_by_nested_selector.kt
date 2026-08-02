// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_by_nested_selector
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val points = listOf(1, 12, 3, 20, 17)
            val grouped = points.groupBy { if (it >= 10) "high" else "low" }
            __check((grouped["high"]!!.size).toString(), "2")
            __check((grouped["low"]!!.size).toString(), "3")
        }
