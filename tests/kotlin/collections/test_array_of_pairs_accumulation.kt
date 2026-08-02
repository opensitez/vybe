// vybe-test: kotlin/collections/test_array_of_pairs_accumulation
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val pairs = arrayOf(Pair(1, 3), Pair(2, 4), Pair(5, 6))
            var sum = 0
            for (item in pairs) {
                sum += item.first + item.second
            }
            println(sum)
            val head = pairs[0]
            println(head.first + head.second)
        }

