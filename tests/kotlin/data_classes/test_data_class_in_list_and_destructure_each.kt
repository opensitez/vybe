// vybe-test: kotlin/data_classes/test_data_class_in_list_and_destructure_each
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class PairNode(val id: Int, val weight: Int)

        fun main() {
            val rows = listOf(PairNode(1, 2), PairNode(3, 4))
            var score = 0
            for ((id, weight) in rows) {
                score += id * weight
            }
            println(score)
        }

