// vybe-test: kotlin/try_catch_flow/test_try_resource_like_sequence
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var opened = 0
            try {
                opened += 1
                try {
                    opened += 10
                } finally {
                    opened += 100
                }
            } finally {
                opened += 1000
            }
            __check((opened).toString(), "1111")
        }
