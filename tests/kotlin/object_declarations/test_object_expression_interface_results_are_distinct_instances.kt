// vybe-test: kotlin/object_declarations/test_object_expression_interface_results_are_distinct_instances
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Counter {
            fun next(): Int
        }

        fun makeCounter(start: Int): Counter {
            var value = start
            return object : Counter {
                override fun next(): Int {
                    value += 1
                    return value
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = makeCounter(0) as Any
            val second = makeCounter(0) as Any
            __check(((first as Counter).next()).toString(), "1")
            __check(((second as Counter).next()).toString(), "1")
            __check((first === second).toString(), "false")
        }
