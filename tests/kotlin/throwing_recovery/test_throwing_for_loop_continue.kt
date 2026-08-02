// vybe-test: kotlin/throwing_recovery/test_throwing_for_loop_continue
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun f(i: Int): Int {
            if (i == 2) throw Exception("bad")
            return i
        }
        fun main() {
            var out = 0
            for (i in 0..3) {
                try {
                    out += f(i)
                } catch (e: Exception) {
                    out += 100
                }
            }
            println(out)
        }

