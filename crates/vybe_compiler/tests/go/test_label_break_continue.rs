//! Labeled break/continue — distinct control transfer out of nested loops.

use crate::helpers::*;

go_run_cases! {
    labeled_break_exits_outer_loop => ("package main; import \"fmt\"; func main() { sum := 0; outer: for i := 0; i < 3; i++ { for j := 0; j < 3; j++ { if i == 1 && j == 1 { break outer }; sum++ } }; fmt.Println(sum) }", vec!["4"]),
    labeled_continue_skips_outer_increment => ("package main; import \"fmt\"; func main() { count := 0; outer: for i := 0; i < 3; i++ { for j := 0; j < 2; j++ { if j == 1 { continue outer }; count++ } }; fmt.Println(count) }", vec!["3"]),
    labeled_break_on_search_found => ("package main; import \"fmt\"; func main() { grid := [][]int{{1,2},{3,4}}; found := -1; search: for r := 0; r < len(grid); r++ { for c := 0; c < len(grid[r]); c++ { if grid[r][c] == 3 { found = r*10 + c; break search } } }; fmt.Println(found) }", vec!["10"]),
    unlabeled_break_inner_only => ("package main; import \"fmt\"; func main() { total := 0; for i := 0; i < 3; i++ { for j := 0; j < 3; j++ { if j == 1 { break }; total++ } }; fmt.Println(total) }", vec!["3"]),
    unlabeled_continue_inner_skip => ("package main; import \"fmt\"; func main() { sum := 0; for i := 0; i < 4; i++ { if i == 2 { continue }; sum += i }; fmt.Println(sum) }", vec!["3"]),
}

go_compile_cases! {
    labeled_break_on_switch_inside_loop => "package main; func main() { loop: for i := 0; i < 2; i++ { switch i { case 1: break loop } } }",
    labeled_continue_on_switch_inside_loop => "package main; func main() { loop: for i := 0; i < 2; i++ { switch i { case 0: continue loop } } }",
    goto_forward_label => "package main; func main() { goto End; End: }",
}
