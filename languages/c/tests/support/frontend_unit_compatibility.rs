pub(super) const UNIT_BOUNDARY_PROGRAMS: &[&str] = &[
    "#include <SDL_messagebox.h>\nint main(void){return SDL_MESSAGEBOX_ERROR!=16 \
     || SDL_MESSAGEBOX_WARNING!=32 || SDL_MESSAGEBOX_INFORMATION!=64 \
     || SDL_MESSAGEBOX_BUTTONS_LEFT_TO_RIGHT!=128 \
     || SDL_MESSAGEBOX_BUTTONS_RIGHT_TO_LEFT!=256 \
     || SDL_MESSAGEBOX_BUTTON_RETURNKEY_DEFAULT!=1 \
     || SDL_MESSAGEBOX_BUTTON_ESCAPEKEY_DEFAULT!=2 \
     || SDL_MESSAGEBOX_COLOR_BACKGROUND!=0 || SDL_MESSAGEBOX_COLOR_TEXT!=1 \
     || SDL_MESSAGEBOX_COLOR_BUTTON_BORDER!=2 || SDL_MESSAGEBOX_COLOR_BUTTON_BACKGROUND!=3 \
     || SDL_MESSAGEBOX_COLOR_BUTTON_SELECTED!=4 || SDL_MESSAGEBOX_COLOR_MAX!=5;}",
    include_str!("../../../../tests/c/doom_startup/frontend_call_handoffs.c"),
    "int f(unsigned n){int s=0;switch(n){case 1:s+=2;case 2:s+=3;break; \
     default:s=7;}return s;}int main(void){return f(1)!=5 || f(2)!=3 || f(0)!=7;}",
    "int main(void){int n=0,i=0;do{n++;continue;n++;}while(0); \
     while(i++<4){if(i==2)continue;n+=i;}do{n++;}while(n<12);return n!=12 || i!=5;}",
    "int route(int x){int a[2]={3,4},b[2]={7,8};int *p;if(x)p=a;else p=b; \
     if(x&&p[0]==3)return p[1];else return p[0];} \
     int main(void){return route(1)!=4 || route(0)!=7;}",
    "static int ticks;int next(void){ticks++;return ticks;}int main(void){int n=0; \
     if(next()==1)if(next()==2)n=7;else n=-1;return n!=7 || ticks!=2;}",
    "typedef unsigned short W; double mix(int a,W b,double c){return a+b+c;} \
     int main(void){return mix(3,4,0.5)!=7.5;}",
    "int main(void){int a=3,b=4; return a+b*2!=11;}",
    "typedef struct Node Node; struct Node {int value; Node *next;}; extern Node *head; \
     int read(void){return head->value+head[1].value;} \
     static Node nodes[2]={{3,0},{4,0}}; Node *head=nodes; \
     int main(void){return read()!=7;}",
    "typedef int (*Callback)(int,...); int choose(int x,...){return x;} \
     static int values[]={3,4}; int invoke(Callback cb){return cb(values[0],values[1])+values[1];} \
     static Callback selected=choose; int main(void){return invoke(selected)!=7;}",
    "typedef unsigned short Word; enum {N=3}; extern const Word table[N]; \
     int total(int n,const Word a[3]){return a[0]+a[n-1];} \
     const Word table[N]={1,2,3}; int main(void){return total(N,table)!=4;}",
    "static int result; void step(int n) { int i=0; while(i<n) { i++; \
     if(i==2) continue; if(i==5) break; result+=i; } if(result==8) goto done; \
     result=-1; done: return; } int main(void) { step(8); return result!=8; }",
    "int total(int n,const int a[3]){return a[0]+a[n-1];} \
     static int table[3]={1,2,3}; int main(void){return total(3,table)!=4;}",
    "int total(int *a[2]){return *a[0]+*a[1];} \
     int main(void){int x=3,y=4; int *a[2]={&x,&y}; return total(a)!=7;}",
    "int total(int (*a)[2]){return a[0][0]+a[1][1];} \
     static int a[2][2]={{1,2},{3,4}}; int main(void){return total(a)!=5;}",
];

pub(super) const USER_CALL_BINDING_PROGRAMS: &[&str] = &[
    "int read(void){return 7;} int main(void){return read()!=7;}",
    "int read(void); int main(void){return read()!=7;} int read(void){return 7;}",
    "int value(void){return 7;} int main(void){int (*read)(void)=value; return read()!=7;}",
    "int value(void){return 7;} int (*write)(void)=value; int main(void){return write()!=7;}",
    "int value(void){return 7;} int invoke(int (*read)(void)){return read();} \
     int main(void){return invoke(value)!=7;}",
    "double value(void){return 1.5;} int invoke(double (*fread)(void),double (*fseek)(void)) \
     {return fread()+fseek()!=3.0;} int main(void){return invoke(value,value);}",
    "extern long read(int,void*,unsigned long); int value(void){return 7;} \
     int invoke(int (*read)(void)){return read();} \
     int main(void){return invoke(value)!=7 || read(0,0,0)!=0;}",
    "extern long read(int,void*,unsigned long); int value(void){return 7;} \
     int main(void){int before=read(0,0,0); {int (*read)(void)=value; \
     if(read()!=7)return 1;} return before || read(0,0,0)!=0;}",
];

pub(super) fn execute_identical_standard_wasm(
    profile_source: &str,
    parsers: [vybe_runtime::registry::ParseFn; 2],
) {
    execute_standard_wasm_programs(
        profile_source,
        parsers,
        UNIT_BOUNDARY_PROGRAMS
            .iter()
            .chain(USER_CALL_BINDING_PROGRAMS)
            .copied(),
    );
}

pub(super) fn execute_standard_wasm_programs<'a>(
    profile_source: &str,
    parsers: [vybe_runtime::registry::ParseFn; 2],
    programs: impl IntoIterator<Item = &'a str>,
) {
    let profile =
        std::sync::Arc::new(vybe_compiler::profile::parse_profile(profile_source).unwrap());
    for source in programs {
        let mut baseline = None;
        let mut ast_baseline = None;
        for parse in parsers {
            let module = parse(source).unwrap();
            let ast = format!("{module:?}");
            if let Some(expected) = &ast_baseline {
                assert_eq!(&ast, expected, "{source}");
            } else {
                ast_baseline = Some(ast);
            }
            let chunks = vybe_compiler::primitives::Compiler::with_shared_profile(profile.clone())
                .compile_with_imports(&module)
                .unwrap()
                .chunks;
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                },
            );
            let mut emission =
                vybe_compiler::primitives::Compiler::with_shared_profile(profile.clone())
                    .begin_emission_owned(module)
                    .unwrap();
            let mut producer = emission.take_owned_units().unwrap();
            let expected_units = producer.remaining_units();
            let mut emitted_units = 0;
            let mut body_started = false;
            while let Some(unit) = producer.next_unit() {
                if unit.is_function_declaration() {
                    assert!(
                        !body_started,
                        "function emitted after source body: {source}"
                    );
                } else {
                    body_started = true;
                }
                let index = unit.source_index();
                let progress = emission.emit_unit(unit).unwrap();
                assert_eq!(progress.source_index, index);
                emitted_units += 1;
            }
            assert!(emitted_units > 0, "{source}");
            assert_eq!(emitted_units, expected_units);
            assert_eq!(producer.remaining_units(), 0);
            assert_eq!(producer.retained_statements(), 0);
            let streamed = emission.finish().unwrap().chunks;
            let streamed_bytes = vybe_platform_wasm::write_wasm_with_options(
                &streamed,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                },
            );
            assert_eq!(bytes, streamed_bytes, "incremental emission: {source}");
            assert_no_dynamic_equality(&streamed, &streamed_bytes);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            assert_no_dynamic_equality(&chunks, &bytes);
            if let Some(expected) = &baseline {
                assert_eq!(&bytes, expected, "{source}");
            } else {
                baseline = Some(bytes);
            }
            for executable in [chunks, streamed] {
                let mut vm = vybe_runtime::VM::new();
                vybe_runtime::init_all_registered(
                    &mut vm,
                    &vybe_runtime::capabilities::Capabilities::all(),
                );
                vm.run(executable).unwrap();
                assert_eq!(vm.pending_exit_code, 0, "{source}");
            }
        }
    }
}

pub(super) fn assert_no_dynamic_equality(chunks: &[vybe_runtime::Chunk], bytes: &[u8]) {
    for chunk in chunks {
        assert!(
            !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
            "{}",
            chunk.name
        );
        let imports = if chunk.imports.is_empty() {
            &chunks[0].imports
        } else {
            &chunk.imports
        };
        let mut offset = 0;
        while offset + 4 <= chunk.code.len() {
            let op = vybe_runtime::opcode::Op::decode(
                chunk.read_u16(offset),
                chunk.read_u16(offset + 2),
            )
            .unwrap();
            if op == vybe_runtime::opcode::Op::CALL {
                if let Some(import) = imports.get(chunk.read_u16(offset + 4) as usize) {
                    assert!(
                        import.module != "ecma:value"
                            || !matches!(import.name.as_str(), "abstractEq" | "abstractNe")
                    );
                }
            }
            let (_, next) = vybe_runtime::debug::disassemble_instruction(chunk, offset);
            assert!(next > offset);
            offset = next;
        }
    }
    let mut abstract_imports = Vec::new();
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    let import = import.unwrap();
                    if matches!(
                        import.ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                    ) {
                        abstract_imports.push(
                            import.module == "ecma:value"
                                && matches!(import.name, "abstractEq" | "abstractNe"),
                        );
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                let mut operators = body.get_operators_reader().unwrap();
                while !operators.eof() {
                    match operators.read().unwrap() {
                        wasmparser::Operator::Call { function_index }
                        | wasmparser::Operator::ReturnCall { function_index }
                        | wasmparser::Operator::RefFunc { function_index } => {
                            assert_ne!(abstract_imports.get(function_index as usize), Some(&true));
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
}
