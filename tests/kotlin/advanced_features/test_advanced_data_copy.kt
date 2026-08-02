// vybe-test: kotlin/advanced_features/test_advanced_data_copy
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

data class Pair(val a: Int, val b: Int)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val p = Pair(1, 2)
val q = p.copy(b = 3)
__check((q.a).toString(), "1")
__check((q.b).toString(), "3") }
