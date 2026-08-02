// vybe-test: kotlin/printing/test_printing_boolean_composition_for_logging
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val active = true
            val ready = false
            __check(("active=${active && !ready}").toString(), "active=true")
            __check(("ready=${!ready}").toString(), "ready=true")
            __check(("both=${active == true && ready == false}").toString(), "both=true")
        }
