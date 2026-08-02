// vybe-test: kotlin/local_functions/test_local_function_accepts_function_argument
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun applyAndDescribe(value: Int, transform: (Int) -> Int): Int {
                return transform(value)
            }
            fun map(v: Int): Int = v + 1
            __check((applyAndDescribe(4, ::map)).toString(), "5")
            __check((applyAndDescribe(4) { it * 2 }).toString(), "8")
        }
