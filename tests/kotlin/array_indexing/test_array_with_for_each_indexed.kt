// vybe-test: kotlin/array_indexing/test_array_with_for_each_indexed
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun main() {
        val a = intArrayOf(1,2,3)
        var out = 0
        a.forEachIndexed { index, value -> if (index == 1) out = value }
        println(out)
    }

