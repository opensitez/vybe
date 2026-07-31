use crate::helpers::run_prints;

#[test]
fn test_measure_time_millis_runs_block() {
    let out = run_prints(r#"
        fun main() {
            var seen = false
            val millis = kotlin.system.measureTimeMillis {
                seen = true
                var sum = 0
                for (i in 1..1000) sum += i
                println(sum)
            }
            println(seen)
            println(millis >= 0)
        }
    "#);
    assert_eq!(out, &["500500", "true", "true"]);
}

#[test]
fn test_measure_nano_time_runs_block() {
    let out = run_prints(r#"
        fun main() {
            var seen = false
            val nanos = kotlin.system.measureNanoTime {
                seen = true
                val out = "x" + "y"
                println(out)
            }
            println(seen)
            println(nanos >= 0)
        }
    "#);
    assert_eq!(out, &["xy", "true", "true"]);
}

#[test]
fn test_measure_time_zero_work_block() {
    let out = run_prints(r#"
        fun main() {
            val value = kotlin.system.measureTimeMillis {
                // no-op
            }
            println(value >= 0)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_measure_nano_time_zero_work_block() {
    let out = run_prints(r#"
        fun main() {
            val value = kotlin.system.measureNanoTime {
                // no-op
            }
            println(value >= 0)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_measure_time_for_loop_scale() {
    let out = run_prints(r#"
        fun main() {
            val tiny = kotlin.system.measureTimeMillis {
                var total = 0
                for (i in 0 until 50000) {
                    total += i
                }
                println(total)
            }
            val larger = kotlin.system.measureTimeMillis {
                var total = 0L
                for (i in 0 until 100000) {
                    total += i
                }
                println(total)
            }
            println(tiny >= 0)
            println(larger >= tiny)
        }
    "#);
    assert_eq!(out, &["1249975000", "4999950000", "true", "true"]);
}

#[test]
fn test_measure_nano_time_for_loop_scale() {
    let out = run_prints(r#"
        fun main() {
            val tiny = kotlin.system.measureNanoTime {
                var total = 0
                for (i in 0 until 20000) {
                    total += i
                }
                println(total)
            }
            val larger = kotlin.system.measureNanoTime {
                var total = 0L
                for (i in 0 until 40000) {
                    total += i.toLong()
                }
                println(total)
            }
            println(tiny >= 0)
            println(larger >= tiny)
        }
    "#);
    assert_eq!(out, &["199990000", "799980000", "true", "true"]);
}

#[test]
fn test_measure_time_nested_blocks() {
    let out = run_prints(r#"
        fun main() {
            val outer = kotlin.system.measureTimeMillis {
                kotlin.system.measureTimeMillis {
                    println("inner")
                }
            }
            println(outer >= 0)
        }
    "#);
    assert_eq!(out, &["inner", "true"]);
}

#[test]
fn test_measure_nano_time_nested_blocks() {
    let out = run_prints(r#"
        fun main() {
            val outer = kotlin.system.measureNanoTime {
                kotlin.system.measureNanoTime {
                    println("inner")
                }
            }
            println(outer >= 0)
        }
    "#);
    assert_eq!(out, &["inner", "true"]);
}

#[test]
fn test_measure_time_with_captured_return_value() {
    let out = run_prints(r#"
        fun main() {
            val result = kotlin.system.measureTimeMillis {
                9 + 1
            }
            val value = result / kotlin.system.measureTimeMillis {
                1
            }
            println(value >= 0)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_measure_nano_time_with_returned_sum() {
    let out = run_prints(r#"
        fun main() {
            val result = kotlin.system.measureNanoTime {
                2 + 3
            }
            println(result >= 0)
            println(kotlin.system.measureTimeMillis {
                7 * 6
            } >= 0)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_measure_time_with_exception_propagation() {
    let out = run_prints(r#"
        fun main() {
            try {
                kotlin.system.measureTimeMillis {
                    throw IllegalStateException("x")
                }
                println("ok")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_measure_nano_time_with_exception_propagation() {
    let out = run_prints(r#"
        fun main() {
            try {
                kotlin.system.measureNanoTime {
                    throw IllegalStateException("n")
                }
                println("ok")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }
    "#);
    assert_eq!(out, &["n"]);
}

#[test]
fn test_measure_time_multiple_invocations_compare_behavior() {
    let out = run_prints(r#"
        fun main() {
            val first = kotlin.system.measureTimeMillis { for (i in 1..1000) {} }
            val second = kotlin.system.measureTimeMillis { for (i in 1..1000) {} }
            println(first >= 0)
            println(second >= 0)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_measure_nano_time_multiple_invocations_compare_behavior() {
    let out = run_prints(r#"
        fun main() {
            val first = kotlin.system.measureNanoTime { for (i in 1..2000) {} }
            val second = kotlin.system.measureNanoTime { for (i in 1..2000) {} }
            println(first >= 0)
            println(second >= 0)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_measure_time_isolated_from_block_state() {
    let out = run_prints(r#"
        fun main() {
            var value = 1
            val elapsed = kotlin.system.measureTimeMillis {
                value *= 2
                println(value)
            }
            println(value)
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["2", "2", "true"]);
}

#[test]
fn test_measure_nano_time_isolated_from_block_state() {
    let out = run_prints(r#"
        fun main() {
            var value = "a"
            val elapsed = kotlin.system.measureNanoTime {
                value += "b"
                println(value)
            }
            println(value)
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["ab", "ab", "true"]);
}

#[test]
fn test_measure_time_vs_nano_time_signals_positive() {
    let out = run_prints(r#"
        fun main() {
            val millis = kotlin.system.measureTimeMillis { for (i in 0 until 3000) {} }
            val nanos = kotlin.system.measureNanoTime { for (i in 0 until 3000) {} }
            println(millis >= 0)
            println(nanos >= 0)
            println((nanos / 1000000) >= millis)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_system_current_time_millis_monotonic_hint() {
    let out = run_prints(r#"
        fun main() {
            val a = System.currentTimeMillis()
            val b = System.currentTimeMillis()
            println(b >= a)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_system_nano_time_changes_over_calls() {
    let out = run_prints(r#"
        fun main() {
            val first = System.nanoTime()
            val second = System.nanoTime()
            println(second >= first)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_system_identity_hash_code_repeatability_for_object() {
    let out = run_prints(r#"
        fun main() {
            val value = arrayOf(1, 2, 3)
            val first = System.identityHashCode(value)
            val second = System.identityHashCode(value)
            println(first != 0)
            println(second != 0)
            println(first == second)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_system_identity_hash_code_differs_for_distinct_objects() {
    let out = run_prints(r#"
        fun main() {
            val a = Any()
            val b = Any()
            val first = System.identityHashCode(a)
            val second = System.identityHashCode(b)
            println(first == second)
            println(first != 0)
            println(second != 0)
        }
    "#);
    assert_eq!(out, &["false", "true", "true"]);
}

#[test]
fn test_measure_time_in_loop_collection_sum() {
    let out = run_prints(r#"
        fun main() {
            val times = mutableListOf<Long>()
            repeat(3) {
                val elapsed = kotlin.system.measureTimeMillis {
                    var sum = 0
                    for (i in 1..5000) sum += i
                    if (sum == 0) println("x")
                }
                times.add(elapsed)
            }
            println(times.size)
            println(times.all { it >= 0 })
        }
    "#);
    assert_eq!(out, &["3", "true"]);
}

#[test]
fn test_measure_nano_time_collection_derived_metric() {
    let out = run_prints(r#"
        fun main() {
            val elapsed = kotlin.system.measureNanoTime {
                val text = listOf("a", "b", "c").joinToString("")
                println(text)
            }
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["abc", "true"]);
}

#[test]
fn test_measure_time_and_runtime_identity_in_one_block() {
    let out = run_prints(r#"
        fun main() {
            val elapsed = kotlin.system.measureTimeMillis {
                val id1 = System.identityHashCode(Any())
                val id2 = System.identityHashCode(Any())
                println(id1 == id2)
            }
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_measure_nano_time_does_not_mask_resulting_output() {
    let out = run_prints(r#"
        fun main() {
            val elapsed = kotlin.system.measureNanoTime {
                val left = kotlin.system.measureTimeMillis { println(1) }
                val right = kotlin.system.measureTimeMillis { println(2) }
                println(left + right)
            }
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["1", "2", "3", "true"]);
}

#[test]
fn test_measure_time_large_block_is_still_non_negative() {
    let out = run_prints(r#"
        fun main() {
            val elapsed = kotlin.system.measureTimeMillis {
                var value = 0
                repeat(10000) {
                    value += it
                }
                println(value)
            }
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["49995000", "true"]);
}

#[test]
fn test_measure_nano_time_large_block_is_still_non_negative() {
    let out = run_prints(r#"
        fun main() {
            val elapsed = kotlin.system.measureNanoTime {
                var value = 0L
                repeat(12000) {
                    value += it.toLong()
                }
                println(value)
            }
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["71994000", "true"]);
}

#[test]
fn test_measure_time_with_sorted_computation() {
    let out = run_prints(r#"
        fun main() {
            val numbers = (1..20).toList().shuffled().sorted()
            val elapsed = kotlin.system.measureTimeMillis {
                println(numbers.joinToString(","))
            }
            println(numbers.size)
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20", "20", "true"]);
}

#[test]
fn test_measure_nano_time_with_map_transformation() {
    let out = run_prints(r#"
        fun main() {
            val items = listOf(1, 2, 3)
            val elapsed = kotlin.system.measureNanoTime {
                val out = items.map { it * 2 }
                println(out.joinToString(","))
            }
            println(elapsed >= 0)
        }
    "#);
    assert_eq!(out, &["2,4,6", "true"]);
}
