// vybe-test: kotlin/try_finally/test_try_finally_with_continue
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun main() {
        var x = 0
        for (i in 1..3) {
            try {
                if (i == 2) continue
                x += 1
            } finally {
                x += 10
            }
        }
        println(x)
    }

