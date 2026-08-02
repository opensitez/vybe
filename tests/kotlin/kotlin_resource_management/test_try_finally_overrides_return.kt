// vybe-test: kotlin/kotlin_resource_management/test_try_finally_overrides_return
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = "init"
            try {
                out = "inside"
                return@main
            } finally {
                out = "final"
                __check((out).toString(), "final")
            }
        }
