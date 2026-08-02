// vybe-test: kotlin/basic/test_modulo_sign_edges
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((-10) % 3).toString(), "-1")
            __check((10 % (-3)).toString(), "-1")
            __check(((-10) % (-3)).toString(), "-1")
        }
