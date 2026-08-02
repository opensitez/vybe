// vybe-test: kotlin/try_finally/test_try_finally_with_break_in_loop
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        for (i in 1..3) {
            try {
                if (i == 2) break
            } finally {
                println(i)
            }
        }
    }

