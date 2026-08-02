// vybe-test: kotlin/this_super/test_this_inside_member_access
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K { val v = 1
fun out() = this.v + 1 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((K().out()).toString(), "2") }
