// vybe-test: kotlin/kotlin_reflection_runtime/test_reflection_when_as_result_type
// origin: languages/kotlin/tests/kotlin/test_kotlin_reflection_runtime.rs

fun main() {
            val values: List<Any> = listOf(Probe("x"), ProbeImpl(1), "str")
            var count = 0
            for (value in values) {
                when {
                    Probe::class.isInstance(value) -> count += 1
                    MarkerContract::class.isInstance(value) -> count += 10
                    else -> count += 100
                }
            }
            println(count)
        }

