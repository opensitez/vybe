use vybe_runtime::VM;

#[test]
fn foreach_protocol_preserves_values_callbacks_and_control_flow() {
    vybe_language_php::register();
    let folder = std::env::temp_dir().join(format!(
        "vybe-foreach-native-{}",
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    std::fs::create_dir(&folder).unwrap();
    let iterator = r#"class Items implements Iterator {
        public $i=0; public $trace='';
        function rewind(){$this->trace.='r';$this->i=0;}
        function valid(){$this->trace.='v';return $this->i<3;}
        function current(){$this->trace.='c';return $this->i+1;}
        function next(){$this->trace.='n';$this->i++;}
        function key(){return $this->i;}
    }
    function total($values){$n=0;foreach($values as $k=>$v){$n+=$v;}return $n;}
    "#;
    let cases = [
        (
            "array_iterator",
            "",
            "if(total(new ArrayIterator([1,2,3]))!=6)exit(1);",
        ),
        ("array", "", "if(total([1,2,3])!=6)exit(1);"),
        (
            "generator",
            "function seq(){yield 1;yield 2;yield 3;}",
            "if(total(seq())!=6)exit(1);",
        ),
        (
            "aggregate",
            "class Group implements IteratorAggregate {function getIterator(){return new Items();}}",
            "if(total(new Group())!=6)exit(1);",
        ),
        (
            "callbacks",
            "function run($it){$n=0;foreach($it as $v){$n+=$v;}return $n;}",
            "$it=new Items();if(run($it)!=6 || $it->trace!='rvcnvcnvcnv')exit(1);",
        ),
        (
            "nested_control",
            "function run($it){$n=0;foreach($it as $v){foreach([1,2,3] as $x){if($x==1)continue;if($v==2)break 2;$n+=$x;}}return $n;}",
            "if(run(new Items())!=5)exit(1);",
        ),
        (
            "throw_finally",
            "function run($it){$n=0;try{foreach($it as $v){if($v==2)throw new Exception('stop');$n+=$v;}}catch(Exception $e){$n+=10;}finally{$n+=100;}return $n;}",
            "if(run(new Items())!=111)exit(1);",
        ),
    ];
    for (name, declarations, check) in cases {
        let path = folder.join(format!("{name}.php"));
        std::fs::write(
            &path,
            format!("<?php\n{iterator}\n{declarations}\n{check}\n"),
        )
        .unwrap();
        let bundle = vybe_compiler::projects::load(&path).unwrap();
        let mut vm = VM::new();
        vybe_runtime::init_all_registered(
            &mut vm,
            &vybe_runtime::capabilities::Capabilities::all(),
        );
        let module = bundle.prepared_module().unwrap();
        let compiled = bundle
            .compile_prepared_module(&module, &vm.modules)
            .unwrap();
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&compiled.chunks))
            .unwrap_or_else(|e| panic!("{name}: invalid Wasm: {e}"));
        vm.run(compiled.chunks)
            .unwrap_or_else(|e| panic!("{name}: {e}"));
        assert_eq!(vm.pending_exit_code, 0, "{name}");
        std::fs::remove_file(path).unwrap();
    }
    std::fs::remove_dir(folder).unwrap();
}
