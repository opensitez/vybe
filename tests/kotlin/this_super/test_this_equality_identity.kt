// vybe-test: kotlin/this_super/test_this_equality_identity
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K { fun same(other: K): Boolean = this === other }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = K()
val b = a
__check((a.same(b)).toString(), "true") }
