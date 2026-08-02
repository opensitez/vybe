// vybe-test: kotlin/this_super/test_this_in_nested_object
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class K {
        private val x = 1
        fun maker() = object {
            fun value() = this@K.x
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((K().maker().value()).toString(), "1") }
