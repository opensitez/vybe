// vybe-test: kotlin/nullability/test_null_elvis_on_boolean
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun main() {
            val flag: Boolean? = null
            val status = flag ?: false
            if (status) {
                println("on")
            } else {
                println("off")
            }
        }

