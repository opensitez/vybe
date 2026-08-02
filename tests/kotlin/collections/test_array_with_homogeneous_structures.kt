// vybe-test: kotlin/collections/test_array_with_homogeneous_structures
// origin: languages/kotlin/tests/kotlin/test_collections.rs

interface Item {
            fun value(): Int
        }

        class NumberItem(val v: Int) : Item {
            override fun value(): Int = v
        }

        fun main() {
            val boxed: Array<Item> = arrayOf(NumberItem(1), NumberItem(2), NumberItem(3))
            var total = 0
            for (item in boxed) {
                total += item.value()
            }
            println(total)
            println(boxed[1].value())
        }

