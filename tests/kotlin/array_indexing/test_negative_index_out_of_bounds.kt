// vybe-test: kotlin/array_indexing/test_negative_index_out_of_bounds
// origin: languages/kotlin/tests/kotlin/test_array_indexing.rs

fun main() { val a = intArrayOf(1, 2)
try { println(a[-1]) } catch (e: Exception) { println("err") } }

