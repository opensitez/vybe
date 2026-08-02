// vybe-test: kotlin/try_finally/test_try_resource_style_manual
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        var data = ""
        try {
            data = "open"
        } finally {
            data = data + "-closed"
        }
        __check((data).toString(), "open-closed")
    }
