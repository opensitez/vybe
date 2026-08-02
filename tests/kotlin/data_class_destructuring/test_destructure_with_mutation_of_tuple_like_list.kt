// vybe-test: kotlin/data_class_destructuring/test_destructure_with_mutation_of_tuple_like_list
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class Entry(val a: Int, val b: Int)

        fun main() {
            val source = listOf(Entry(1, 2), Entry(3, 4))
            var sum = 0
            for ((left, right) in source) {
                sum += left + right
            }
            println(sum)
        }

