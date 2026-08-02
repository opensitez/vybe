// vybe-test: kotlin/kotlin_visibility_advanced/test_private_setter_with_external_caller
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

class Bucket {
            var total = 0
                private set

            fun add(v: Int) {
                total += v
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Bucket()
            b.add(3)
            b.add(4)
            __check((b.total).toString(), "7")
        }
