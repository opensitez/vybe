use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(simple_switch_match, "int x=2; int r=0; switch(x){case 1:r=10;break;case 2:r=20;break;default:r=-1;} System.out.println(r);", "20");
jt!(simple_switch_default, "int x=5; int r=0; switch(x){case 1:r=10;break;case 2:r=20;break;default:r=30;} System.out.println(r);", "30");
jt!(fallthrough_collect, "int x=1; int r=0; switch(x){case 1:r+=1;case 2:r+=2;break;case 3:r+=3;break;default:r+=9;} System.out.println(r);", "3");
jt!(fallthrough_with_break, "int x=2; int r=0; switch(x){case 1:r=1;break;case 2:r=2;break;case 3:r=3;break;default:r=0;} System.out.println(r);", "2");
jt!(switch_with_char, r#"char c='b'; String out=""; switch(c){case 'a':out="A";break;case 'b':out="B";break;default:out="Z";} System.out.println(out);"#, "B");
jt!(string_switch_match, r#"String s="two"; String out=""; switch(s){case "one":out="1";break;case "two":out="2";break;default:out="0";} System.out.println(out);"#, "2");
jt!(string_switch_default, r#"String s="x"; String out=""; switch(s){case "a":out="A";break;case "b":out="B";break;default:out="D";} System.out.println(out);"#, "D");
jt!(switch_expression_in_case, "int x = 2 + 1; int r=0; switch(x){case 1:r=1;break;case 2:r=2;break;case 3:r=3;break;default:r=9;} System.out.println(r);", "3");
jt!(switch_multiple_cases, "int x=4; int r=0; switch(x){case 1:case 2:r=1;break;case 3:case 4:r=2;break;default:r=0;} System.out.println(r);", "2");
jt!(switch_nested_ternary, "int x = 1; int y = 0; switch(x){case 1:y = x==1 ? 11 : 12;break;default:y=0;} System.out.println(y);", "11");
jt!(switch_inside_loop, "int sum=0; for(int x=1; x<=3; x++){switch(x){case 1:sum+=1;break;case 2:sum+=2;break;default:sum+=10;}} System.out.println(sum);", "13");
jt!(switch_on_modulo, "int x=7; int r=0; switch(x % 3){case 1:r=10;break;case 2:r=20;break;default:r=30;} System.out.println(r);", "10");
jt!(switch_case_no_break_cascade, "int x=1; int r=0; switch(x){case 1:r=1;case 2:r+=2;case 3:r+=3;break;default:r+=9;} System.out.println(r);", "6");
jt!(switch_all_cases_true, "int x=3; int r=0; switch(x){case 1:r=1;break;case 2:r=2;break;case 3:r=3;break;} System.out.println(r);", "3");
jt!(switch_default_first, "int x=9; int r=0; switch(x){default:r=7;break;case 1:r=1;} System.out.println(r);", "7");
jt!(switch_after_declaration, "int x=1; int y=0; if (x==1){switch(x){case 1:y=1;break;default:y=2;}} System.out.println(y);", "1");
jt!(switch_in_nested_block, "int x=2; int y=0; if (x>0){switch(x){case 1:y=1;break;case 2:y=2;break;default:y=3;}} System.out.println(y);", "2");
jt!(switch_boolean_simulated, "int x=(1>0)?1:0; int y=0; switch(x){case 0:y=0;break;case 1:y=1;break;} System.out.println(y);", "1");
jt!(switch_with_post_increment, r#"int x=1; String label; switch(x++){case 1: label="one"; break; default: label="no";} System.out.println(label + "|" + x);"#, "one|2");
jt!(switch_on_byte, "byte x=1; int y=0; switch(x){case 1:y=4;break;case 2:y=8;break;default:y=0;} System.out.println(y);", "4");
jt!(switch_with_negative, r#"int x=-1; String out=""; switch(x){case -1:out="neg";break;case 1:out="pos";break;default:out="zero";} System.out.println(out);"#, "neg");
jt!(switch_string_builder_concat, r#"String x="ok"; StringBuilder sb = new StringBuilder(); switch(x){case "ok":sb.append("yes");break;default:sb.append("no");} System.out.println(sb.toString());"#, "yes");
jt!(switch_on_enforced_bounds, "int x=0; int r=0; switch(x){case 0:r=100;break;case 1:r=200;break;default:r=300;} System.out.println(r);", "100");
jt!(switch_with_returning_value, r#"int x=4; String out; switch(x){case 1:out="a";break;case 2:out="b";break;case 3:out="c";break;default:out="d";} System.out.println(out);"#, "d");
jt!(switch_uses_constant_expr, "final int target=2; int r=0; switch(2 + 1){case target:r=9;break;case 3:r=12;break;default:r=0;} System.out.println(r);", "12");
