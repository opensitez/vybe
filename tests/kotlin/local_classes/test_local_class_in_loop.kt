// vybe-test: kotlin/local_classes/test_local_class_in_loop
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun main() {
            val out = StringBuilder()
            for (i in 1..3) {
                class Local(val v: Int)
                out.append(Local(i).v)
            }
            println(out.toString())
        }

