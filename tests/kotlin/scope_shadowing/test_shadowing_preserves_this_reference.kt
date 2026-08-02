// vybe-test: kotlin/scope_shadowing/test_shadowing_preserves_this_reference
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

class Box {
            val value = "box"
            fun run(): String {
                val value = "inner"
                return this.value
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().run()).toString(), "box")
        }
