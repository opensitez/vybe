// vybe-test: kotlin/preconditions/test_precondition_chain_for_password
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun validate(password: String?) {
            requireNotNull(password)
            require(password.length >= 4, { "short" })
            require(password.any { it.isDigit() }, { "digit missing" })
        }

        fun main() {
            try {
                validate("a1")
                println("ok")
            } catch (e: IllegalArgumentException) {
                println(e.message)
            }
        }

