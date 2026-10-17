const PROGRAMS: &[&str] = &[
    include_str!("../../../tests/c/doom_startup/nested_label_goto.c"),
    "int main(void){int n=0;goto right;if(1){left:n+=3;goto done;}\
     else{right:n+=5;goto left;}done:return n!=8;}",
    "int main(void){int n=0;start:n++;if(n<3)goto deep;return n!=3;\
     if(0){deep:goto start;}}",
    "int main(void){int i=0,checks=0,total=0;goto body;while(++checks<3){\
     total+=100;body:total++;if(++i==1)continue;total+=10;break;}\
     return total!=112 || checks!=1;}",
    "int main(void){int n=0,checks=0;goto body;do{n+=100;body:n++;}\
     while(++checks<2);return n!=102 || checks!=2;}",
    "int main(void){int init=0,checks=0,updates=0,total=0;goto body;\
     for(init++;++checks<3;updates++){total+=100;body:total++;\
     if(total==1)continue;break;}return init!=0 || checks!=1 || updates!=1 || total!=102;}",
    "static int calls;int probe(void){calls++;return 1;}\
     int main(void){int n=0;goto body;switch(probe()){case 1:n=100;\
     body:n+=3;break;default:n=-1;}n+=4;return n!=7 || calls!=0;}",
    "static int calls;int probe(void){calls++;return 7;}\
     int main(void){int n=0;goto body;{int skipped=probe();\
     body:n+=sizeof(skipped)==sizeof(int);}return n!=1 || calls!=0;}",
    "int main(void){int n=0,result=0;again:{int local=++n;\
     if(local<3)goto again;result=local;}return n!=3 || result!=3;}",
    "int f(void){int n=0;goto inside;if(0){inside:n=3;}return n;}\
     int g(void){int n=0;goto inside;if(0){inside:n=4;}return n;}\
     int main(void){return f()!=3 || g()!=4;}",
    "int route(int enter){int n=0;if(enter)goto body;switch(0){default:n+=2;\
     body:n+=3;case 1:n+=4;}return n;}int main(void){return route(1)!=7 || route(0)!=9;}",
    "int main(void){int i=0,n=0;goto body;for(;i<3;i++){switch(i){\
     case 0:body:n++;continue;default:n+=10;break;}n+=100;}return i!=3 || n!=221;}",
];

#[test]
fn nested_c_labels_lower_at_function_scope_without_residual_gotos() {
    for source in PROGRAMS {
        let expected = format!("{:?}", crate::goto_lowering_c_tests::parse_prepared_peg(source).unwrap());
        assert_eq!(format!("{:?}", crate::goto_lowering_c_tests::parse_prepared_native(source).unwrap()), expected, "{source}");
        assert!(!expected.contains("GoTo("), "unlowered C goto: {source}");
    }
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn nested_c_label_entries_preserve_effects_continuations_and_standard_wasm_execution() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::goto_lowering_c_tests::parse_prepared_peg,
         crate::goto_lowering_c_tests::parse_prepared_native], PROGRAMS.iter().copied());
}
