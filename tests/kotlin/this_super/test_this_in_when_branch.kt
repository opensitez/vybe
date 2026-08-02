// vybe-test: kotlin/this_super/test_this_in_when_branch
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K { fun kind(v: Int): String = when (v) { 1 -> this.javaClass.simpleName
else -> "n" } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((K().kind(1)).toString(), "K") }
