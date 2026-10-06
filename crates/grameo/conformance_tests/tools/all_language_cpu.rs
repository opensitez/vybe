//! Matched in-memory parser CPU; reference derives are supplied from frozen grammars.
use std::{env, fs, hint::black_box, io::{BufReader, Read}};
use pest::Parser as _;
use vybe_parser::{MatchOptions, program::Program};
include!(env!("VYBE_CPU_REFERENCES"));
mod python_helper { include!(env!("VYBE_AUDIT_PYTHON_HELPER")); }
#[allow(clippy::collapsible_if)]
mod fortran_helper { include!(env!("VYBE_AUDIT_FORTRAN_HELPER")); }
unsafe extern "C" { fn js_thread_cpu_ns() -> u64; }
fn cpu() -> u64 { let value=unsafe { js_thread_cpu_ns() }; assert_ne!(value,u64::MAX); value }
fn number(r: &mut impl Read) -> u64 { let mut b=[0;8];r.read_exact(&mut b).unwrap();u64::from_le_bytes(b) }
fn string(r: &mut impl Read) -> String { let n=number(r);let mut b=vec![0;n as usize];r.read_exact(&mut b).unwrap();String::from_utf8(b).unwrap() }
type Shape=Vec<(String,usize,usize,usize)>;
fn pest_shape<R:pest::RuleType>(pairs:pest::iterators::Pairs<'_,R>) -> (usize,Shape) {
    fn visit<R:pest::RuleType>(p:pest::iterators::Pair<'_,R>,d:usize,s:&mut Shape) { s.push((format!("{:?}",p.as_rule()),p.as_span().start(),p.as_span().end(),d));for c in p.into_inner(){visit(c,d+1,s);} }
    let mut s=vec![];let mut consumed=0;for p in pairs { consumed=p.as_span().end();visit(p,0,&mut s); }(consumed,s)
}
fn native_shape(tree:vybe_parser::ParseTree<'_,'_>) -> (usize,Shape) {
    fn visit(p:vybe_parser::Pair<'_,'_,'_>,d:usize,s:&mut Shape) {let span=p.as_span();s.push((p.rule_name().into(),span.start,span.end,d));for c in p.into_inner(){visit(c,d+1,s);}}
    let mut s=vec![];for p in tree.pairs(){visit(p,0,&mut s);} (tree.consumed(),s)
}
fn fingerprint(label:&str,source:&str)->u64 {let mut v=0xcbf29ce484222325u64;for p in [label.as_bytes(),source.as_bytes()]{for b in (p.len() as u64).to_le_bytes().iter().chain(p){v=(v^u64::from(*b)).wrapping_mul(0x100000001b3);}}v}
struct Case {label:String,source:String,accepted:bool}
// Keep every backend behind the same opaque timing callsite. Specialized
// closure loops produced different timings even for identical parser calls.
#[inline(never)]
fn timed(f:&dyn Fn(),n:usize)->f64 {let f=black_box(f);let n=black_box(n);f();let t=cpu();for _ in 0..n {f();} (cpu()-t) as f64/n as f64}
fn iterations(f:&dyn Fn())->usize {for _ in 0..4{f();}let ns=timed(f,4);(250_000.0/ns.max(1.0)).clamp(1.0,128.0) as usize}
fn run<R:pest::RuleType,P:Program>(language:&str,rule:R,pest:impl for<'a> Fn(R,&'a str)->Result<pest::iterators::Pairs<'a,R>,pest::error::Error<R>>,generated:&P,entry:&str,args:&[String]) {
    let resolved=vybe_parser::compile(&fs::read_to_string(format!("{}/{language}/src/grammar.pest",args[2])).unwrap()).unwrap();
    let mut reader=BufReader::new(fs::File::open(format!("{}/{language}.inputs",args[1])).unwrap());
    let options=MatchOptions::default();let mut cases=vec![];
    for _ in 0..number(&mut reader) {let label=string(&mut reader);let raw=string(&mut reader);let transform=number(&mut reader);let bytes=number(&mut reader);let hash=number(&mut reader);let accepted=number(&mut reader)!=0;
        let source=if transform!=0 && language=="python" {python_helper::preprocess_indentation(&raw)} else if transform!=0 && language=="fortran" {fortran_helper::rewrite_hollerith(&raw)} else {raw};
        assert_eq!(source.len() as u64,bytes);assert_eq!(fingerprint(&label,&source),hash);
        fs::write(format!("{}/matched-{language}-{}.source",args[1],cases.len()),&source).unwrap();
        let reference=pest(rule,&source).map(pest_shape);let native=vybe_parser::engine::parse_program(generated,entry,&source,options).map(native_shape);let dynamic=vybe_parser::engine::parse_program(&resolved,entry,&source,options).map(native_shape);
        assert_eq!(reference.is_ok(),accepted,"{language}:{label}");assert_eq!(native.is_ok(),accepted);assert_eq!(dynamic.is_ok(),accepted);
        if let Ok(expected)=reference {assert_eq!(native.unwrap(),expected,"generated:{language}:{label}");assert_eq!(dynamic.unwrap(),expected,"resolved:{language}:{label}");}
        else {assert_eq!(native.unwrap_err().kind,vybe_parser::ParseErrorKind::Syntax);assert_eq!(dynamic.unwrap_err().kind,vybe_parser::ParseErrorKind::Syntax);}
        println!("PARITY\t{language}\t{label}\t{bytes}\t{accepted}");cases.push(Case{label,source,accepted});
    }
    let mut counts=vec![];
    for c in &cases {let p=iterations(&||{drop(black_box(pest(rule,black_box(&c.source))));});let g=iterations(&||{drop(black_box(vybe_parser::engine::parse_program(generated,entry,black_box(&c.source),options)));});let d=iterations(&||{drop(black_box(vybe_parser::engine::parse_program(&resolved,entry,black_box(&c.source),options)));});counts.push(p.max(g).max(d));}
    for round in 0..args[3].parse::<usize>().unwrap() {for (i,c) in cases.iter().enumerate(){
        let p=||{drop(black_box(pest(rule,black_box(&c.source))));};
        let g=||{drop(black_box(vybe_parser::engine::parse_program(generated,entry,black_box(&c.source),options)));};
        let d=||{drop(black_box(vybe_parser::engine::parse_program(&resolved,entry,black_box(&c.source),options)));};
        // Cover both directions of every ordering, then give each callback the
        // same immediate history: one untimed parse of its own workload.
        let orders=[[0,1,2],[1,2,0],[2,0,1],[2,1,0],[1,0,2],[0,2,1]];
        for backend in orders[(round+i)%orders.len()] {let n=counts[i];let action:&dyn Fn()=match backend {0=>&p,1=>&g,_=>&d};let ns=timed(action,n);println!("TIME\t{language}\t{}\t{}\t{round}\t{backend}\t{n}\t{ns:.3}",c.label,c.accepted);}
    }}
}
fn main(){std::thread::Builder::new().stack_size(64*1024*1024).spawn(||{let args=env::args().collect::<Vec<_>>();assert_eq!(args.len(),4);run_selected(&args);}).unwrap().join().unwrap();}
