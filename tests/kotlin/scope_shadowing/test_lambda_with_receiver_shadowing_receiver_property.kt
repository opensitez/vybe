// vybe-test: kotlin/scope_shadowing/test_lambda_with_receiver_shadowing_receiver_property
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

class Profile(val name: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name = "outer"
            val profile = Profile("inner")
            val out = profile.run {
                val name = this.name
                val innerName = "inner2"
                innerName
            }
            __check((out).toString(), "inner2")
            __check((name).toString(), "outer")
        }
