// vybe-test: kotlin/spread_arguments/test_spread_boolean_vararg
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun allTrue(vararg values: Boolean): Boolean {
            for (v in values) if (!v) return false
            return true
        }
        fun main() {
            val flags = booleanArrayOf(true, true, false)
            println(allTrue(*flags))
            val flags2 = booleanArrayOf(true, true)
            println(allTrue(*flags2))
        }

