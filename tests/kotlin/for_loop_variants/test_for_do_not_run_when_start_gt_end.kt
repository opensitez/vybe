// vybe-test: kotlin/for_loop_variants/test_for_do_not_run_when_start_gt_end
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var out = 0
            for (i in 5 downTo 8) {
                out += i
            }
            println(out)
        }

