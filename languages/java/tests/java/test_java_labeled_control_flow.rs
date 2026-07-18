use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(simple_label_break, "int total = 0; outer: for (int i = 0; i < 5; i++) { for (int j = 0; j < 5; j++) { if (j == 2) break outer; total += 1; } } System.out.println(total);", "2");
jt!(simple_label_continue, "int total = 0; outer: for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 1) continue outer; total += 1; } } System.out.println(total);", "3");
jt!(label_break_with_nested_if, "int a = 0; int i = 0; outer: while (i < 5) { if (i == 3) break outer; a += i; i++; } System.out.println(a);", "3");
jt!(label_continue_while, "int a = 0; int i = 0; outer: while (i < 5) { i++; if (i % 2 == 0) continue outer; a++; } System.out.println(a);", "3");
jt!(multiple_labels, "int a=0; first: for(int i=0;i<3;i++){ second: for(int j=0;j<3;j++){ a+=1; if(j==1) continue second; } } System.out.println(a);", "6");
jt!(label_between_loops, "int total = 0; outer: for(int i=0;i<3;i++){ for(int j=0;j<3;j++){ if(i==1) { continue outer; } total += 1; } } System.out.println(total);", "6");
jt!(label_break_inner_only, "int total = 0; for(int i=0;i<3;i++){ for(int j=0;j<3;j++){ if(j==1) break; total += 1; } } System.out.println(total);", "6");
jt!(label_unused_definition, "int a = 0; label: if (a == 0) { a = 1; } System.out.println(a);", "1");
jt!(label_before_while_loop, "int a = 0; int i = 0; outer: while(i < 4){ if(i == 2) { i++; continue outer; } a += i; i++; } System.out.println(a);", "2");
jt!(label_break_out_of_if, "int a = 0; int i = 0; outer: for(; i < 4; i++) { for(int j = 0; j < 4; j++) { if (i == 2) break outer; a++; } } System.out.println(a);", "6");
jt!(label_continue_in_for, "int a = 0; outer: for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 0) continue outer; a++; } } System.out.println(a);", "0");
jt!(label_continue_from_for_each, "int[] arr = {1,2,3,4}; int s = 0; outer: for (int v : arr) { if (v == 2) continue outer; s += v; } System.out.println(s);", "8");
jt!(label_break_from_do, "int i = 0; int s = 0; outer: do { i++; if (i == 2) break outer; s += i; } while (i < 5); System.out.println(s);", "1");
jt!(label_continue_in_do, "int i = 0; int s = 0; outer: do { i++; if (i % 2 == 0) continue outer; s++; } while (i < 5); System.out.println(s);", "3");
jt!(label_nested_switch, "int total = 0; int x = 0; outer: switch(x){ case 0: for(int i=0;i<3;i++){ if(i==1) break outer; total += i; } break; default: total = 9; } System.out.println(total);", "0");
jt!(label_for_continue_with_else, "int a = 0; int b = 0; outer: for (int i = 0; i < 4; i++) { if (i == 2) { continue outer; } b += i; } System.out.println(b);", "4");
jt!(label_multi_level_break, "int i = 0; int c = 0; outer: for (; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 1) break outer; c++; } } System.out.println(c);", "0");
jt!(label_break_in_while_switch, "int c = 0; int i = 0; outer: while (true) { if (i == 2) break outer; c += i; i++; } System.out.println(c);", "1");
jt!(label_continue_across_nested, "int c = 0; int i = 0; outer: for (; i < 4; i++) { for (int j = 0; j < 4; j++) { if (j == 2) continue outer; c++; } } System.out.println(c);", "8");
jt!(label_break_while_nested, "int i = 0; int j = 0; int c = 0; outer: while (i < 3) { i++; while (j < 3) { if (j == 1) break outer; c++; j++; } } System.out.println(c);", "0");
jt!(label_continue_while_nested, "int i = 0; int c = 0; outer: for (i = 0; i < 3; i++) { int j = 0; while (j < 3) { j++; if (j == 1) continue outer; c++; } } System.out.println(c);", "6");
jt!(label_break_with_variable, "int a = 0; int b = 0; outer: for (int i = 0; i < 4; i++) { for (int j = 0; j < 4; j++) { b++; if (i == 1 && j == 1) break outer; } a++; } System.out.println(a + "," + b);", "1,2");
jt!(label_continue_after_calc, "int s = 0; outer: for (int i = 0; i < 6; i++) { if (i % 2 == 0) continue outer; s += i; } System.out.println(s);", "9");
jt!(label_break_if_condition, "int x=0; outer: for(int i = 0; i < 5; i++) { if (i == 3) break outer; x += i; } System.out.println(x);", "3");
jt!(label_continue_in_block, "int x=0; outer: for(int i = 0; i < 5; i++) { { if (i == 3) continue outer; } x += i; } System.out.println(x);", "7");
jt!(label_after_for_each, "int[] values = {1,2,3,4}; int s = 0; outer: for (int v : values) { if (v == 1) continue outer; if (v == 4) break outer; s += v; } System.out.println(s);", "5");
jt!(label_nested_three, "int c=0; outer: for(int i=0;i<3;i++){ for(int j=0;j<3;j++){ for(int k=0;k<3;k++){ if (k == 1) continue outer; c++; } } } System.out.println(c);", "9");
jt!(label_break_three, "int c=0; outer: for(int i=0;i<3;i++){ for(int j=0;j<3;j++){ for(int k=0;k<3;k++){ if (j == 1) break outer; c++; } } } System.out.println(c);", "3");
