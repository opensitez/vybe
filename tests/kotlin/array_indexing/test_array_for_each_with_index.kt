// vybe-test: kotlin/array_indexing/test_array_for_each_with_index
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun main() {
        val a = intArrayOf(1, 2, 3)
        var sum = 0
        for (i in a.indices) { sum += a[i] }
        println(sum)
    }

