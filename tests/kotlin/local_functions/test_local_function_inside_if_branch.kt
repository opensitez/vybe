// vybe-test: kotlin/local_functions/test_local_function_inside_if_branch
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val isAdmin = true
            val value = if (isAdmin) {
                fun role(): String = "admin"
                role()
            } else {
                fun role(): String = "user"
                role()
            }
            __check((value).toString(), "admin")
        }
