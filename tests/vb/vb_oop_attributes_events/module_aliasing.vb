' vybe-test: vb/vb_oop_attributes_events/module_aliasing
' origin: languages/vb/tests/vb/test_vb_oop_attributes_events.rs

Imports Alias = System.Console: Module M: Sub Main(): Alias.WriteLine("Alias"): End Sub: End Module
