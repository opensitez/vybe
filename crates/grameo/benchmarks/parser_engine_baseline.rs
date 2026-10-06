use std::{hint::black_box,time::Instant};
use pest::Parser as _;
#[allow(dead_code,clippy::upper_case_acronyms)] mod old_lua { include!("lua_before_generated.rs"); }
#[allow(dead_code,clippy::upper_case_acronyms)] mod new_lua { include!("lua_generated.rs"); }
#[allow(dead_code,clippy::upper_case_acronyms)] mod old_js { include!("js_before_generated.rs"); }
#[allow(dead_code,clippy::upper_case_acronyms)] mod new_js { include!("js_generated.rs"); }
mod lua_reference { use pest_derive::Parser; #[derive(Parser)] #[grammar="examples/lua.pest"] pub struct Parser; }
mod js_reference { use pest_derive::Parser; #[derive(Parser)] #[grammar="examples/js.pest"] pub struct Parser; }
type Shape=Vec<(String,usize,usize,usize)>;
fn old_shape(tree:vybe_parser_before::ParseTree<'_, '_>)->(usize,Shape) {
 fn visit(pair:vybe_parser_before::Pair<'_, '_,'_>, depth:usize,out:&mut Shape) { let span=pair.as_span();out.push((pair.rule_name().to_owned(),span.start,span.end,depth));for child in pair.into_inner(){visit(child,depth+1,out)} }
 let mut shape=Vec::new();for pair in tree.pairs(){visit(pair,0,&mut shape)};(tree.consumed(),shape)
}
fn new_shape(tree:vybe_parser::ParseTree<'_, '_>)->(usize,Shape) {
 fn visit(pair:vybe_parser::Pair<'_, '_,'_>, depth:usize,out:&mut Shape) { let span=pair.as_span();out.push((pair.rule_name().to_owned(),span.start,span.end,depth));for child in pair.into_inner(){visit(child,depth+1,out)} }
 let mut shape=Vec::new();for pair in tree.pairs(){visit(pair,0,&mut shape)};(tree.consumed(),shape)
}
fn pest_shape<R:pest::RuleType>(pairs:pest::iterators::Pairs<'_,R>)->(usize,Shape) {
 fn visit<R:pest::RuleType>(pair:pest::iterators::Pair<'_,R>,depth:usize,out:&mut Shape){let span=pair.as_span();out.push((format!("{:?}",pair.as_rule()),span.start(),span.end(),depth));for child in pair.into_inner(){visit(child,depth+1,out)}}
 let mut shape=Vec::new();let mut end=0;for pair in pairs {end=pair.as_span().end();visit(pair,0,&mut shape)};(end,shape)
}
fn main(){
 let lua=std::env::args().any(|a|a=="--lua");let count=std::env::args().find_map(|a|a.parse::<usize>().ok()).unwrap_or(501).max(1);
 let mut cases:Vec<(&str,String)> = if lua {vec![
 ("tiny","return 1 + 2 * 3".into()),
 ("declarations100",(0..100).map(|n|format!("local x{n} = {n}\n")).collect()),
 ("expressions100",(0..100).map(|n|format!("local x{n} = f(1 + 2 * 3, (4 - 5) / 6)\n")).collect()),
 ("comments100",(0..100).map(|n|format!("-- comment {n}\nlocal x{n} = {n}\n")).collect())
 ]}else{vec![
 ("tiny","const x = 1 + 2 * 3;".into()),
 ("declarations100",(0..100).map(|n|format!("const x{n} = {n};\n")).collect()),
 ("expressions100",(0..100).map(|n|format!("const x{n} = f(1 + 2 * 3, (4 - 5) / 6);\n")).collect()),
 ("comments100",(0..100).map(|n|format!("// comment {n}\nconst x{n} = {n};\n")).collect())]};
 let entry=if lua{"chunk"}else{"program"};
 for (_,source) in &cases {
 let old=if lua{old_lua::Parser.parse(old_lua::Rule::chunk,source)}else{old_js::Parser.parse(old_js::Rule::program,source)}.map(old_shape).unwrap();
 let new=if lua{new_lua::Parser.parse(new_lua::Rule::chunk,source)}else{new_js::Parser.parse(new_js::Rule::program,source)}.map(new_shape).unwrap();
 let reference=if lua{pest_shape(lua_reference::Parser::parse(lua_reference::Rule::chunk,source).unwrap())}else{pest_shape(js_reference::Parser::parse(js_reference::Rule::program,source).unwrap())};assert_eq!(old,new);assert_eq!(new,reference);
 }
 for (_,source) in &cases {
 for input in [source.clone(),format!("{source}\0"),source[..source.len()/2].to_owned()] {
 for depth in [0,1,2,4,8,16] {for budget in [0,1,2,8,32,128,1024,8192,65536] {
 let old_options=vybe_parser_before::MatchOptions{max_steps:budget,max_rule_depth:depth};
 let new_options=vybe_parser::MatchOptions{max_steps:budget,max_rule_depth:depth};
 let old=if lua{vybe_parser_before::engine::parse_program(&old_lua::Parser,entry,&input,old_options)}else{vybe_parser_before::engine::parse_program(&old_js::Parser,entry,&input,old_options)}.map(old_shape).map_err(|e|format!("{e:?}"));
 let new=if lua{vybe_parser::engine::parse_program(&new_lua::Parser,entry,&input,new_options)}else{vybe_parser::engine::parse_program(&new_js::Parser,entry,&input,new_options)}.map(new_shape).map_err(|e|format!("{e:?}"));assert_eq!(old,new,"{entry} {depth} {budget}");
 }}}
 }
 println!("Full original/candidate/pest captures match; 648 old/new resource comparisons pass.");
 for (label,source) in cases.drain(..){
 let mut times:[Vec<f64>;3]=std::array::from_fn(|_|Vec::with_capacity(count));
 for round in 0..count+10 {for i in 0..3 {let variant=(i+round)%3;let now=Instant::now();
 match variant {
 0=>{if lua{black_box(old_lua::Parser.parse(old_lua::Rule::chunk,black_box(&source)).unwrap());}else{black_box(old_js::Parser.parse(old_js::Rule::program,black_box(&source)).unwrap());}},
 1=>{if lua{black_box(new_lua::Parser.parse(new_lua::Rule::chunk,black_box(&source)).unwrap());}else{black_box(new_js::Parser.parse(new_js::Rule::program,black_box(&source)).unwrap());}},
 _=>{if lua{black_box(lua_reference::Parser::parse(lua_reference::Rule::chunk,black_box(&source)).unwrap());}else{black_box(js_reference::Parser::parse(js_reference::Rule::program,black_box(&source)).unwrap());}}
 }
 if round>=10{times[variant].push(now.elapsed().as_secs_f64()*1e6)}
 }}
 let stats=times.map(|mut t|{t.sort_by(f64::total_cmp);(t[count/2],t[count*95/100])});println!("{label}: {} bytes; original {:?}; candidate {:?}; pest {:?}",source.len(),stats[0],stats[1],stats[2]);
 }
}
