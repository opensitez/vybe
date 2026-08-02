// vybe-test: kotlin/function_overloads/test_overload_top_level_and_local_same_name
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun ping(v: Int): String = "global"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun ping(v: String): String = "local"
            __check((ping(1)).toString(), "global")
            __check((ping("a")).toString(), "local")
        }
