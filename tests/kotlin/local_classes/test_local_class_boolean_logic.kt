// vybe-test: kotlin/local_classes/test_local_class_boolean_logic
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Gate(val open: Boolean)
            __check((Gate(true).open).toString(), "true")
            __check((Gate(false).open).toString(), "false")
        }
