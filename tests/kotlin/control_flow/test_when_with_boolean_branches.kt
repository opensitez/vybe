// vybe-test: kotlin/control_flow/test_when_with_boolean_branches
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            val isReady = true
            when {
                isReady == false -> println("no")
                isReady && true -> println("yes")
                else -> println("maybe")
            }
        }

