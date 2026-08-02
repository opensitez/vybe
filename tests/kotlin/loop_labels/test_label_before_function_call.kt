// vybe-test: kotlin/loop_labels/test_label_before_function_call
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun tick(): Int = 1
        fun main() {
            var out = 0
            loop@ while (true) {
                if (tick() == 0) continue
                out += 1
                if (out == 4) break@loop
            }
            println(out)
        }

