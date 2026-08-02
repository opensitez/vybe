// vybe-test: kotlin/this_super/test_this_is_used_in_comparison
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K { val v = 1
fun same(other: K) = this.v == other.v }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((K().same(K())).toString(), "true") }
