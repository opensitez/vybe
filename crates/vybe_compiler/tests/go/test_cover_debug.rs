//! debug/* binary formats — one API per test.


go_compile_cases! {
    debug_buildinfo_read => "package main; import \"runtime/debug\"; func main() { _, _ = debug.ReadBuildInfo() }",
    debug_buildinfo_module_version => "package main; import \"runtime/debug\"; func main() { info, _ := debug.ReadBuildInfo(); if info != nil { _ = info.Main.Version } }",
    dwarf_new => "package main; import \"debug/dwarf\"; import \"bytes\"; func main() { _, _ = dwarf.New(bytes.NewReader(nil)) }",
    dwarf_tag_string => "package main; import \"debug/dwarf\"; func main() { _ = dwarf.Tag(0).String() }",
    dwarf_attr_string => "package main; import \"debug/dwarf\"; func main() { _ = dwarf.Attr(0).String() }",
    elf_new_file => "package main; import \"debug/elf\"; import \"bytes\"; func main() { _, _ = elf.NewFile(bytes.NewReader(nil)) }",
    elf_class_string => "package main; import \"debug/elf\"; func main() { _ = elf.Class(0).String() }",
    elf_machine_string => "package main; import \"debug/elf\"; func main() { _ = elf.Machine(0).String() }",
    macho_new_file => "package main; import \"debug/macho\"; import \"bytes\"; func main() { _, _ = macho.NewFile(bytes.NewReader(nil)) }",
    macho_cpu_string => "package main; import \"debug/macho\"; func main() { _ = macho.Cpu(0).String() }",
    pe_new_file => "package main; import \"debug/pe\"; import \"bytes\"; func main() { _, _ = pe.NewFile(bytes.NewReader(nil)) }",
    pe_machine_string => "package main; import \"debug/pe\"; func main() { _ = pe.Machine(0).String() }",
    plan9obj_new_file => "package main; import \"debug/plan9obj\"; import \"bytes\"; func main() { _, _ = plan9obj.NewFile(bytes.NewReader(nil)) }",
    gosym_new_table => "package main; import \"debug/gosym\"; import \"bytes\"; func main() { _, _ = gosym.NewTable(nil, bytes.NewReader(nil)) }",
    gosym_sym_kind_string => "package main; import \"debug/gosym\"; func main() { _ = gosym.SymKind(0).String() }",
}
