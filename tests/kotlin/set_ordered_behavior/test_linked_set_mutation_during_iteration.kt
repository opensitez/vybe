// vybe-test: kotlin/set_ordered_behavior/test_linked_set_mutation_during_iteration
// origin: languages/kotlin/tests/kotlin/test_set_ordered_behavior.rs

fun main() {
            val values = linkedSetOf(1, 2, 3)
            val outValues = StringBuilder()
            val it = values.iterator()
            while (it.hasNext()) {
                val n = it.next()
                if (n == 2) {
                    it.remove()
                }
                outValues.append(n)
            }
            println(outValues.toString())
            println(values.joinToString(","))
        }

