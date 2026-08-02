// vybe-test: kotlin/repeat_statements/test_repeat_returnless
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun repeatAndAdd(start: Int): Int {
            var out = start
            repeat(3) {
                out += 2
            }
            return out
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((repeatAndAdd(1)).toString(), "7")
        }
