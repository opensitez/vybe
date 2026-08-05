use crate::helpers::run_prints;

#[test]
fn test_sequence_from_list_is_lazy_until_terminal() {
    let out = run_prints(
        r#"
        fun main() {
            var built = 0
            val source = listOf(1, 2, 3)
            val seq = source.asSequence().onEach { built += 1 }
            println("before")
            println(seq.count())
            println("after")
            println(built)
        }
    "#,
    );
    assert_eq!(out, &["before", "3", "after", "3"]);
}

#[test]
fn test_sequence_to_list_materializes_values() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = sequenceOf(1, 2, 3).map { it * 10 }
            val out = seq.toList()
            println(out.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["10,20,30"]);
}

#[test]
fn test_sequence_filter_map_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = (1..6).asSequence()
                .map { it + 1 }
                .filter { it % 2 == 0 }
            println(seq.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6"]);
}

#[test]
fn test_sequence_take_and_drop() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = (1..10).asSequence()
            println(seq.take(4).toList().joinToString(","))
            println((1..10).asSequence().drop(7).toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4", "8,9,10"]);
}

#[test]
fn test_sequence_take_while_drop_while() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(1, 3, 5, 2, 4, 6).asSequence()
            println(seq.takeWhile { it < 4 }.toList().joinToString(","))
            println(seq.dropWhile { it < 4 }.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3", "5,2,4,6"]);
}

#[test]
fn test_generate_sequence_with_seed_and_limit() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = generateSequence(1) { v -> if (v < 4) v + 1 else null }
            println(seq.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_generate_sequence_string_builder() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = generateSequence("a") { prev -> if (prev.length < 3) prev + prev else null }
            println(seq.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,aa,aaaa"]);
}

#[test]
fn test_sequence_distinct_values() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(1, 1, 2, 2, 3).asSequence().distinct()
            println(seq.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_sequence_distinct_by_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf("aa", "ab", "b", "cc").asSequence().distinctBy { it.length }
            println(seq.toList().joinToString("|"))
        }
    "#,
    );
    // The separator the code passes is `"|"` — distinct lengths keep "aa"
    // (2) and "b" (1), so the joined output is `aa|b`.
    assert_eq!(out, &["aa|b"]);
}

#[test]
fn test_sequence_zip_with_collection() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = (1..4).asSequence().zip(listOf("a", "b", "c", "d")) { n, s -> "$n-$s" }
            println(seq.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1-a,2-b,3-c,4-d"]);
}

#[test]
fn test_sequence_flat_map_nested_lists() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(listOf(1, 2), listOf(3, 4)).asSequence().flatMap { it.asSequence() }
            println(seq.sum())
            println(seq.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["10", "1,2,3,4"]);
}

#[test]
fn test_sequence_flatten_words() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(listOf("a"), emptyList(), listOf("b", "c")).asSequence().flatten()
            println(seq.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,b,c"]);
}

#[test]
fn test_sequence_join_to_string() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = (1..4).asSequence()
            println(seq.joinToString(prefix = "[", postfix = "]"))
        }
    "#,
    );
    assert_eq!(out, &["[1, 2, 3, 4]"]);
}

#[test]
fn test_sequence_fold_reduce() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(1, 2, 3, 4).asSequence()
            println(seq.fold(0) { acc, n -> acc + n })
            println(listOf(4, 3, 2).asSequence().reduce { acc, n -> acc - n })
        }
    "#,
    );
    assert_eq!(out, &["10", "-1"]);
}

#[test]
fn test_sequence_any_all_none() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(2, 4, 6).asSequence()
            println(seq.any { it > 5 })
            println(seq.all { it % 2 == 0 })
            println(seq.none { it == 9 })
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_sequence_count_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(1, 2, 3, 4, 5, 6).asSequence()
            println(seq.count { it % 3 == 0 })
            println(seq.count())
        }
    "#,
    );
    assert_eq!(out, &["2", "6"]);
}

#[test]
fn test_sequence_sorted_and_reversed() {
    let out = run_prints(
        r#"
        fun main() {
            val source = listOf("pear", "apple", "kiwi").asSequence()
            println(source.sorted().joinToString(","))
            println(source.sortedDescending().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["apple,kiwi,pear", "pear,kiwi,apple"]);
}

#[test]
fn test_sequence_sorted_by_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val source = listOf("aa", "b", "ccc").asSequence()
            println(source.sortedBy { it.length }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["b,aa,ccc"]);
}

#[test]
fn test_sequence_group_by_and_count() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf("cat", "dog", "cow", "deer").asSequence()
            val grouped = seq.groupBy { it.first() }
            val c = grouped["c"]?.size ?: 0
            val d = grouped["d"]?.size ?: 0
            println(c)
            println(d)
        }
    "#,
    );
    assert_eq!(out, &["2", "2"]);
}

#[test]
fn test_sequence_partition() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = (1..6).asSequence()
            val (lt4, ge4) = seq.partition { it < 4 }
            println(lt4.joinToString(","))
            println(ge4.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "4,5,6"]);
}

#[test]
fn test_sequence_with_chunked_windows() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = (1..5).asSequence()
            println(seq.chunked(2).joinToString("|") { it.joinToString("-") })
            println((1..5).asSequence().windowed(3).joinToString("|") { it.joinToString("-") })
        }
    "#,
    );
    assert_eq!(out, &["1-2|3-4|5", "1-2-3|2-3-4|3-4-5"]);
}

#[test]
fn test_sequence_single_and_single_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val seqOne = listOf(7).asSequence()
            val seqEmpty = emptySequence<Int>()
            println(seqOne.single())
            println(seqEmpty.singleOrNull() ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["7", "-1"]);
}

#[test]
fn test_sequence_find_first_last() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(5, 12, 17, 20).asSequence()
            println(seq.find { it % 2 == 0 } ?: -1)
            println(seq.findLast { it % 2 == 1 } ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["12", "17"]);
}

#[test]
fn test_sequence_first_or_null_last_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(9, 10, 11).asSequence()
            println(seq.firstOrNull { it > 100 } ?: "none")
            println(seq.lastOrNull { it < 10 } ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["none", "9"]);
}

#[test]
fn test_sequence_iterator_contract() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = sequenceOf(1, 2)
            val it = seq.iterator()
            println(it.hasNext())
            println(it.next())
            println(it.hasNext())
            println(it.next())
            println(it.hasNext())
        }
    "#,
    );
    assert_eq!(out, &["true", "1", "true", "2", "false"]);
}

#[test]
fn test_sequence_to_set_and_to_list_stability() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf(1, 2, 2, 3).asSequence()
            val asSet = seq.toSet()
            val asList = seq.toList()
            println(asSet.joinToString(","))
            println(asList.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "4"]);
}

#[test]
fn test_empty_sequence_has_expected_behaviors() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = emptySequence<Int>()
            println(seq.count())
            println(seq.none { true })
            println(seq.toList().size)
        }
    "#,
    );
    assert_eq!(out, &["0", "true", "0"]);
}

#[test]
fn test_sequence_element_at_or_null_and_last() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = listOf("x", "y", "z").asSequence()
            println(seq.elementAt(1))
            println(seq.elementAtOrNull(5) ?: "none")
            println(seq.last())
        }
    "#,
    );
    assert_eq!(out, &["y", "none", "z"]);
}

#[test]
fn test_sequence_map_with_stateful_side_effect() {
    let out = run_prints(
        r#"
        fun main() {
            var seen = 0
            val seq = (1..5).asSequence().map { n ->
                seen += 1
                n * 10
            }
            println("start")
            println(seq.take(3).toList().joinToString(","))
            println(seen)
            println(seq.toList().size)
            println(seen)
        }
    "#,
    );
    assert_eq!(out, &["start", "10,20,30", "3", "2", "5"]);
}

#[test]
fn test_sequence_with_generator_and_take() {
    let out = run_prints(
        r#"
        fun main() {
            var calls = 0
            val seq = sequence {
                var x = 0
                while (x < 5) {
                    yield(x)
                    calls += 1
                    x += 1
                }
            }
            println(seq.take(3).toList().joinToString(","))
            println(calls)
        }
    "#,
    );
    assert_eq!(out, &["0,1,2", "3"]);
}

#[test]
fn test_sequence_reused_iterable_runs_again_when_materialized_twice() {
    let out = run_prints(
        r#"
        fun main() {
            var calls = 0
            val seq = sequence {
                calls += 1
                yield(1)
                calls += 1
                yield(2)
            }
            println(seq.toList().joinToString(","))
            println(seq.toList().joinToString(","))
            println(calls)
        }
    "#,
    );
    assert_eq!(out, &["1,2", "1,2", "4"]);
}

#[test]
fn test_sequence_empty_first_throws_without_default() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println(emptySequence<Int>().first())
            } catch (e: NoSuchElementException) {
                println("empty")
            }
        }
    "#,
    );
    assert_eq!(out, &["empty"]);
}

#[test]
fn test_sequence_for_each_executes_for_all_elements_and_maintains_order() {
    let out = run_prints(
        r#"
        fun main() {
            var seen = ""
            (1..4).asSequence().forEach { seen += it.toString() }
            println(seen)
        }
    "#,
    );
    assert_eq!(out, &["1234"]);
}

#[test]
fn test_sequence_map_indexed_projects_index_and_value() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = sequenceOf("a", "bb", "ccc")
            val annotated = seq.mapIndexed { index, value ->
                value.length + index
            }.toList().joinToString(",")
            println(annotated)
        }
    "#,
    );
    assert_eq!(out, &["1,3,5"]);
}

#[test]
fn test_sequence_filter_not_null_skips_nulls() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = sequenceOf(1, null, 2, null, 3).filterNotNull()
            println(seq.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_sequence_zip_with_next_returns_adjacent_pairs() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = (1..5).asSequence().zipWithNext { a, b -> a * 10 + b }
            println(seq.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["12,23,34,45"]);
}

#[test]
fn test_sequence_running_fold_and_reduce_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = sequenceOf(1, 2, 3, 4).runningFold(0) { acc, value -> acc + value }.toList()
            println(seq.joinToString(","))
            val running = sequenceOf(1, 2, 3).runningReduce { acc, value -> acc * value }.toList()
            println(running.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,1,3,6,10", "1,2,6"]);
}

#[test]
fn test_sequence_to_sorted_set_uses_sorted_order() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = sequenceOf(3, 1, 4, 1, 5, 2).toSortedSet()
            println(seq.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4,5"]);
}

#[test]
fn test_sequence_generate_sequence_side_effects_are_short_circuited_by_take() {
    let out = run_prints(
        r#"
        fun main() {
            var calls = 0
            val seq = sequence {
                yield(1)
                calls += 1
                yield(2)
                calls += 1
                yield(3)
                calls += 1
                yield(4)
            }
            println(seq.take(3).toList().joinToString(","))
            println(calls)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "2"]);
}

#[test]
fn test_sequence_constrain_once_prevents_reuse() {
    let out = run_prints(
        r#"
        fun main() {
            val constrained = sequenceOf(1, 2, 3).constrainOnce()
            println(constrained.toList().joinToString(","))
            try {
                println(constrained.toList().joinToString(","))
            } catch (e: IllegalStateException) {
                println("cannot_reuse")
            }
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "cannot_reuse"]);
}

#[test]
fn test_sequence_empty_reduce_throws() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println((emptySequence<Int>()).reduce { acc, value -> acc + value })
            } catch (e: Exception) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["error"]);
}

#[test]
fn test_sequence_take_last_and_drop_last() {
    let out = run_prints(
        r#"
        fun main() {
            val source = (1..5).asSequence()
            println(source.takeLast(3).toList().joinToString(","))
            println(source.dropLast(1).toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,4,5", "1,2,3,4"]);
}

#[test]
fn test_sequence_windowed_with_partial_windows() {
    let out = run_prints(
        r#"
        fun main() {
            val source = (1..5).asSequence()
            println(source.windowed(2, 2, partialWindows = true).joinToString("|") { it.joinToString("-") })
            println(source.windowed(3).toList().isEmpty())
        }
    "#,
    );
    assert_eq!(out, &["1-2|3-4|5", "false"]);
}

#[test]
fn test_sequence_zip_stops_at_shortest_input() {
    let out = run_prints(
        r#"
        fun main() {
            val zipped = (1..5).asSequence().zip(listOf("a", "b", "c")) { n, s -> "$n-$s" }
            println(zipped.toList().joinToString(","))
            println(zipped.count())
        }
    "#,
    );
    assert_eq!(out, &["1-a,2-b,3-c", "3"]);
}

#[test]
fn test_sequence_filter_indexed_uses_index_semantics() {
    let out = run_prints(
        r#"
        fun main() {
            val out = (10..16).asSequence()
                .filterIndexed { index, value -> index % 2 == 0 && value > 11 }
                .toList()
                .joinToString(",")
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["12,14,16"]);
}

#[test]
fn test_sequence_map_indexed_keeps_index_contract() {
    let out = run_prints(
        r#"
        fun main() {
            val labeled = listOf("x", "y", "z").asSequence()
                .mapIndexed { index, value -> "$index:$value" }
                .toList()
            println(labeled.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0:x,1:y,2:z"]);
}

#[test]
fn test_sequence_element_at_or_else_throws_and_recovers() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println((1..3).asSequence().elementAt(5))
            } catch (e: IndexOutOfBoundsException) {
                println("oor")
            }
            println((1..3).asSequence().elementAtOrNull(5) ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["oor", "none"]);
}
