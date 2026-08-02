// vybe-test: kotlin/loop_labels/test_label_break_in_repeat
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var i = 0
            var out = 0
            repeat@ repeat(5) {
                i += 1
                if (i == 3) {
                    out = 99
                    return@repeat
                }
            }
            __check((out).toString(), "99")
        }
