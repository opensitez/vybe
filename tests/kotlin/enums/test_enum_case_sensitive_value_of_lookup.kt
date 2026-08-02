// vybe-test: kotlin/enums/test_enum_case_sensitive_value_of_lookup
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Flag { Enabled, Disabled }

        fun main() {
            try {
                println(Flag.valueOf("Enabled"))
            } catch (e: Exception) {
                println("bad")
            }

            try {
                println(Flag.valueOf("enabled"))
                println("should not happen")
            } catch (e: Exception) {
                println("missing")
            }
        }

