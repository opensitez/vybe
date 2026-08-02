// vybe-test: kotlin/this_super/test_this_in_anonymous_object
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val o = object {
            val value = 9
            fun read(): Int = this.value
        }
        __check((o.read()).toString(), "9")
    }
