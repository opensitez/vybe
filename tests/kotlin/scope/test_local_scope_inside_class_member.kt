// vybe-test: kotlin/scope/test_local_scope_inside_class_member
// origin: languages/kotlin/tests/kotlin/test_scope.rs

class Box {
            var value = 0
            fun addStep(step: Int): Int {
                val value = step
                fun bump(): Int {
                    return this.value + value
                }
                this.value = step * 2
                return bump()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.addStep(4)).toString(), "4")
            __check((b.value).toString(), "8")
        }
