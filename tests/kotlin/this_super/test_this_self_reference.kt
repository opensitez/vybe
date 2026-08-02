// vybe-test: kotlin/this_super/test_this_self_reference
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { class K { fun id(): String = this.toString() }
__check((K().id().isNotEmpty()).toString(), "true") }
