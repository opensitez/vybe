// vybe-test: kotlin/control_flow/test_repeat_iterates_exact_times
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            repeat(4) { index ->
                total += index
            }
            __check((total).toString(), "6")
            var markers = ""
            repeat(0) {
                markers += "x"
            }
            __check((markers.isEmpty()).toString(), "true")
        }
