//! Shared loop scaffold changes must preserve PHP control transfers.
use vybe_runtime::VM;

#[test]
fn for_loop_control_transfers_preserve_results() {
    vybe_language_php::register();
    let cases = [
        ("plain", "for($i=0;$i<10;$i++){$sum+=$i;}", 45.0),
        ("continue", "for($i=0;$i<10;$i++){if($i%2===0)continue;$sum+=$i;}", 25.0),
        ("break", "for($i=0;$i<10;$i++){if($i===5)break;$sum+=$i;}", 10.0),
        ("finally", "for($i=0;$i<5;$i++){try{if($i===2)continue;$sum+=$i;}finally{$sum+=10;}}", 58.0),
        ("nested", "for($i=0;$i<3;$i++){$j=0;while($j<3){$j++;if($j===2)continue;$sum++;}}", 6.0),
        ("continue_update", "for($i=0;$i<3;$i++){if($i===1)continue;$sum=$sum*10+$i;}", 2.0),
        ("while_true", "$i=0;while(true){$sum+=$i;if(++$i===5)break;}", 10.0),
        ("for_true", "for($i=0;true;$i++){if($i===5)break;$sum+=$i;}", 10.0),
        ("while_false", "while(false){$sum=99;}", 0.0),
        ("for_false", "for($i=0;false;$i++){$sum=99;}", 0.0),
        ("do_false_continue", "do{$sum++;continue;}while(false);", 1.0),
        ("do_true", "do{$sum++;if($sum===5)break;}while(true);", 5.0),
        ("do_true_continue", "do{$sum++;if($sum<3)continue;break;}while(true);", 3.0),
    ];
    for (name, body, expected) in cases {
        let module = vybe_language_php::parse(&format!("<?php $sum=0;{body}return $sum;")).unwrap();
        let profile = vybe_compiler::profile::parse_profile(
            vybe_language_php::profile_source(),
        ).unwrap();
        let chunks = vybe_compiler::Compiler::with_profile(profile)
            .compile(&module).unwrap();
        let wasm = vybe_platform_wasm::write_wasm_with_options(&chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            });
        vybe_platform_wasm::validate_binary(&wasm)
            .unwrap_or_else(|error| panic!("{name}: {error}"));
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        let result = vm.run(chunks).unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(result.as_f64(), expected, "{name}");
    }
}
