// vybe-test: kotlin/scope_shadowing/test_destructure_local_overwrites
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "outer"
            val (value, count) = listOf("x", "y").withIndex().first()
            __check((value).toString(), "x")
            __check((count).toString(), "0")
            __check(("outer").toString(), "outer")
        }
