use super::helpers::run_vb;

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, vec![$($expected),*]);
        }
    };
}

vb_full_spec!(namespace_spec_module_member_can_be_called_without_instance, r#"Module Util
    Public Function Name() As String
        Return "util"
    End Function
End Module
Module M
    Sub Main()
        Console.WriteLine(Util.Name())
    End Sub
End Module"#, ["util"]);
vb_full_spec!(namespace_spec_module_function_can_be_qualified_by_module_name, r#"Module TextUtil
    Public Function Upper(value As String) As String
        Return UCase(value)
    End Function
End Module
Module M
    Sub Main()
        Console.WriteLine(TextUtil.Upper("vb"))
    End Sub
End Module"#, ["VB"]);
vb_full_spec!(namespace_spec_module_sub_can_mutate_module_field, r#"Module Counter
    Public Value As Integer
    Public Sub Increment()
        Value += 1
    End Sub
End Module
Module M
    Sub Main()
        Counter.Increment()
        Console.WriteLine(Counter.Value)
    End Sub
End Module"#, ["1"]);
vb_full_spec!(namespace_spec_module_const_can_be_read, r#"Module Config
    Public Const Name As String = "vybe"
End Module
Module M
    Sub Main()
        Console.WriteLine(Config.Name)
    End Sub
End Module"#, ["vybe"]);
vb_full_spec!(namespace_spec_module_property_round_trips_value, r#"Module Config
    Public Property Name As String
End Module
Module M
    Sub Main()
        Config.Name = "vb"
        Console.WriteLine(Config.Name)
    End Sub
End Module"#, ["vb"]);
vb_full_spec!(namespace_spec_nested_namespace_class_can_be_instantiated, r#"Namespace Demo.Core
    Public Class Person
        Public Function Name() As String
            Return "person"
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Core.Person()).Name())
    End Sub
End Module"#, ["person"]);
vb_full_spec!(namespace_spec_nested_namespace_module_function_can_be_called, r#"Namespace Demo.Core
    Public Module Util
        Public Function Name() As String
            Return "util"
        End Function
    End Module
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.Core.Util.Name())
    End Sub
End Module"#, ["util"]);
vb_full_spec!(namespace_spec_imports_namespace_enables_short_type_name, r#"Namespace Demo.Core
    Public Class Person
        Public Function Name() As String
            Return "person"
        End Function
    End Class
End Namespace
Imports Demo.Core
Module M
    Sub Main()
        Console.WriteLine((New Person()).Name())
    End Sub
End Module"#, ["person"]);
vb_full_spec!(namespace_spec_imports_alias_can_target_namespace, r#"Namespace Demo.Core
    Public Class Person
        Public Function Name() As String
            Return "person"
        End Function
    End Class
End Namespace
Imports CoreAlias = Demo.Core
Module M
    Sub Main()
        Console.WriteLine((New CoreAlias.Person()).Name())
    End Sub
End Module"#, ["person"]);
vb_full_spec!(namespace_spec_imports_alias_can_target_type, r#"Namespace Demo.Core
    Public Class Person
        Public Function Name() As String
            Return "person"
        End Function
    End Class
End Namespace
Imports PersonAlias = Demo.Core.Person
Module M
    Sub Main()
        Console.WriteLine((New PersonAlias()).Name())
    End Sub
End Module"#, ["person"]);
vb_full_spec!(namespace_spec_fully_qualified_name_can_disambiguate_types, r#"Namespace LeftSpace
    Public Class Token
        Public Function Name() As String
            Return "left"
        End Function
    End Class
End Namespace
Namespace RightSpace
    Public Class Token
        Public Function Name() As String
            Return "right"
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New LeftSpace.Token()).Name())
        Console.WriteLine((New RightSpace.Token()).Name())
    End Sub
End Module"#, ["left", "right"]);
vb_full_spec!(namespace_spec_two_namespaces_can_define_same_class_name, r#"Namespace A
    Public Class ValueBox
        Public Function Name() As String
            Return "A"
        End Function
    End Class
End Namespace
Namespace B
    Public Class ValueBox
        Public Function Name() As String
            Return "B"
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New A.ValueBox()).Name())
        Console.WriteLine((New B.ValueBox()).Name())
    End Sub
End Module"#, ["A", "B"]);
vb_full_spec!(namespace_spec_module_inside_namespace_can_hold_function, r#"Namespace Demo
    Public Module MathUtil
        Public Function DoubleValue(x As Integer) As Integer
            Return x * 2
        End Function
    End Module
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.MathUtil.DoubleValue(8))
    End Sub
End Module"#, ["16"]);
vb_full_spec!(namespace_spec_class_inside_namespace_can_reference_sibling_class, r#"Namespace Demo
    Public Class NameBox
        Public Shared Function Value() As Integer
            Return Counter.NextValue()
        End Function
    End Class
    Public Class Counter
        Private Shared currentValue As Integer = 0
        Public Shared Function NextValue() As Integer
            currentValue += 1
            Return currentValue
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.NameBox.Value())
        Console.WriteLine(Demo.NameBox.Value())
    End Sub
End Module"#, ["1", "2"]);
vb_full_spec!(namespace_spec_module_method_can_return_array, r#"Module Data
    Public Function Build() As Integer()
        Return New Integer() {1, 2, 3}
    End Function
End Module
Module M
    Sub Main()
        Console.WriteLine(Data.Build()(1))
    End Sub
End Module"#, ["2"]);
vb_full_spec!(namespace_spec_module_can_contain_enum, r#"Module Data
    Public Enum Tone
        Low
        High
    End Enum
End Module
Module M
    Sub Main()
        Console.WriteLine(Data.Tone.High)
    End Sub
End Module"#, ["1"]);
vb_full_spec!(namespace_spec_module_can_contain_structure, r#"Module Data
    Public Structure Point
        Public X As Integer
    End Structure
End Module
Module M
    Sub Main()
        Dim p As Data.Point
        p.X = 7
        Console.WriteLine(p.X)
    End Sub
End Module"#, ["7"]);
vb_full_spec!(namespace_spec_module_can_contain_class, r#"Module Data
    Public Class Box
        Public Value As String = "box"
    End Class
End Module
Module M
    Sub Main()
        Console.WriteLine((New Data.Box()).Value)
    End Sub
End Module"#, ["box"]);
vb_full_spec!(namespace_spec_imports_multiple_namespaces_can_support_two_types, r#"Namespace Demo.A
    Public Class BoxA
        Public Function Name() As String
            Return "A"
        End Function
    End Class
End Namespace
Namespace Demo.B
    Public Class BoxB
        Public Function Name() As String
            Return "B"
        End Function
    End Class
End Namespace
Imports Demo.A
Imports Demo.B
Module M
    Sub Main()
        Console.WriteLine((New BoxA()).Name())
        Console.WriteLine((New BoxB()).Name())
    End Sub
End Module"#, ["A", "B"]);
vb_full_spec!(namespace_spec_namespace_alias_can_select_specific_type, r#"Namespace Demo.Core
    Public Class Widget
        Public Function Name() As String
            Return "widget"
        End Function
    End Class
End Namespace
Imports W = Demo.Core.Widget
Module M
    Sub Main()
        Console.WriteLine((New W()).Name())
    End Sub
End Module"#, ["widget"]);
vb_full_spec!(namespace_spec_nested_module_name_can_be_qualified, r#"Namespace Demo.Core
    Public Module Names
        Public Function Value() As String
            Return "names"
        End Function
    End Module
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.Core.Names.Value())
    End Sub
End Module"#, ["names"]);
vb_full_spec!(namespace_spec_partial_namespace_segments_can_resolve_type, r#"Namespace Demo
    Namespace Core
        Public Class Token
            Public Function Name() As String
                Return "token"
            End Function
        End Class
    End Namespace
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Core.Token()).Name())
    End Sub
End Module"#, ["token"]);
vb_full_spec!(namespace_spec_module_function_can_be_imported_by_short_name, r#"Namespace Demo
    Public Module Util
        Public Function Value() As String
            Return "ok"
        End Function
    End Module
End Namespace
Imports Demo.Util
Module M
    Sub Main()
        Console.WriteLine(Value())
    End Sub
End Module"#, ["ok"]);
vb_full_spec!(namespace_spec_class_inside_namespace_can_inherit_base_in_same_namespace, r#"Namespace Demo
    Public Class Animal
        Public Overridable Function Speak() As String
            Return "animal"
        End Function
    End Class
    Public Class Dog
        Inherits Animal
        Public Overrides Function Speak() As String
            Return "dog"
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Dog()).Speak())
    End Sub
End Module"#, ["dog"]);
vb_full_spec!(namespace_spec_class_inside_namespace_can_implement_interface_in_same_namespace, r#"Namespace Demo
    Public Interface INameable
        Function Name() As String
    End Interface
    Public Class Widget
        Implements INameable
        Public Function Name() As String Implements INameable.Name
            Return "widget"
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Dim value As Demo.INameable = New Demo.Widget()
        Console.WriteLine(value.Name())
    End Sub
End Module"#, ["widget"]);
vb_full_spec!(namespace_spec_namespace_can_contain_generic_class, r#"Namespace Demo
    Public Class Box(Of T)
        Public Value As T
        Public Sub New(v As T)
            Value = v
        End Sub
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Box(Of Integer)(9)).Value)
    End Sub
End Module"#, ["9"]);
vb_full_spec!(namespace_spec_module_can_call_class_in_same_namespace, r#"Namespace Demo
    Public Class Box
        Public Value As String = "box"
    End Class
    Public Module Util
        Public Function Read() As String
            Return (New Box()).Value
        End Function
    End Module
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.Util.Read())
    End Sub
End Module"#, ["box"]);
vb_full_spec!(namespace_spec_imported_namespace_can_expose_nested_type, r#"Namespace Demo.Core
    Public Class Box
        Public Function Name() As String
            Return "box"
        End Function
    End Class
End Namespace
Imports Demo.Core
Module M
    Sub Main()
        Console.WriteLine((New Box()).Name())
    End Sub
End Module"#, ["box"]);
vb_full_spec!(namespace_spec_alias_import_can_coexist_with_full_name, r#"Namespace Demo.Core
    Public Class Box
        Public Function Name() As String
            Return "box"
        End Function
    End Class
End Namespace
Imports CoreAlias = Demo.Core
Module M
    Sub Main()
        Console.WriteLine((New CoreAlias.Box()).Name())
        Console.WriteLine((New Demo.Core.Box()).Name())
    End Sub
End Module"#, ["box", "box"]);
vb_full_spec!(namespace_spec_root_level_module_and_namespace_types_can_coexist, r#"Module Util
    Public Function RootName() As String
        Return "root"
    End Function
End Module
Namespace Demo
    Public Class Box
        Public Function Name() As String
            Return Util.RootName()
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Box()).Name())
    End Sub
End Module"#, ["root"]);
vb_full_spec!(namespace_spec_namespace_type_can_use_module_helper_in_same_namespace, r#"Namespace Demo
    Public Module Util
        Public Function GetName() As String
            Return "demo"
        End Function
    End Module
    Public Class Box
        Public Function Name() As String
            Return Util.GetName()
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Box()).Name())
    End Sub
End Module"#, ["demo"]);
vb_full_spec!(namespace_spec_nested_namespace_can_contain_enum, r#"Namespace Demo.Core
    Public Enum Tone
        Low
        High
    End Enum
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.Core.Tone.High)
    End Sub
End Module"#, ["1"]);
vb_full_spec!(namespace_spec_imports_alias_can_shorten_nested_namespace_chain, r#"Namespace Demo.Core.Tools
    Public Class Box
        Public Function Name() As String
            Return "tools"
        End Function
    End Class
End Namespace
Imports ToolsAlias = Demo.Core.Tools
Module M
    Sub Main()
        Console.WriteLine((New ToolsAlias.Box()).Name())
    End Sub
End Module"#, ["tools"]);
vb_full_spec!(namespace_spec_module_sub_can_be_called_from_class_method, r#"Namespace Demo
    Public Module Util
        Public Function Value() As String
            Return "util"
        End Function
    End Module
    Public Class Box
        Public Function Read() As String
            Return Util.Value()
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Box()).Read())
    End Sub
End Module"#, ["util"]);
vb_full_spec!(namespace_spec_class_method_can_call_module_function_in_same_namespace, r#"Namespace Demo
    Public Module Util
        Public Function Value() As Integer
            Return 7
        End Function
    End Module
    Public Class Box
        Public Function Read() As Integer
            Return Util.Value()
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New Demo.Box()).Read())
    End Sub
End Module"#, ["7"]);
vb_full_spec!(namespace_spec_namespace_qualified_enum_value_can_be_printed, r#"Namespace Demo
    Public Enum Size
        Small = 1
        Large = 2
    End Enum
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.Size.Large)
    End Sub
End Module"#, ["2"]);
vb_full_spec!(namespace_spec_namespace_qualified_structure_value_can_be_used, r#"Namespace Demo
    Public Structure Point
        Public X As Integer
    End Structure
End Namespace
Module M
    Sub Main()
        Dim p As Demo.Point
        p.X = 6
        Console.WriteLine(p.X)
    End Sub
End Module"#, ["6"]);
vb_full_spec!(namespace_spec_module_property_can_be_accessed_with_namespace_qualification, r#"Namespace Demo
    Public Module Config
        Public Property Name As String
    End Module
End Namespace
Module M
    Sub Main()
        Demo.Config.Name = "cfg"
        Console.WriteLine(Demo.Config.Name)
    End Sub
End Module"#, ["cfg"]);
vb_full_spec!(namespace_spec_alias_import_can_create_generic_list, r#"Imports IntList = System.Collections.Generic.List(Of Integer)
Module M
    Sub Main()
        Dim items As New IntList()
        items.Add(3)
        Console.WriteLine(items(0))
    End Sub
End Module"#, ["3"]);
vb_full_spec!(namespace_spec_namespace_type_can_reference_global_system_type, r#"Namespace Demo
    Public Class Clock
        Public Function NowText() As String
            Return Global.System.DateTime.Now.Year.ToString()
        End Function
    End Class
End Namespace
Module M
    Sub Main()
        Console.WriteLine(IsNumeric((New Demo.Clock()).NowText()))
    End Sub
End Module"#, ["true"]);

vb_full_spec!(namespace_spec_module_can_shadow_imported_name_with_local_variable, r#"Namespace Demo
    Public Module Util
        Public Function Value() As String
            Return "module"
        End Function
    End Module
End Namespace
Imports Demo
Module M
    Sub Main()
        Dim Util As String = "local"
        Console.WriteLine(Util)
    End Sub
End Module"#, ["local"]);
vb_full_spec!(namespace_spec_module_can_call_sibling_module_function, r#"Namespace Demo
    Public Module A
        Public Function Name() As String
            Return B.Name()
        End Function
    End Module
    Public Module B
        Public Function Name() As String
            Return "b"
        End Function
    End Module
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.A.Name())
    End Sub
End Module"#, ["b"]);
vb_full_spec!(namespace_spec_namespace_can_contain_multiple_modules, r#"Namespace Demo
    Public Module A
        Public Function LeftValue() As Integer
            Return 2
        End Function
    End Module
    Public Module B
        Public Function RightValue() As Integer
            Return 5
        End Function
    End Module
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.A.LeftValue() + Demo.B.RightValue())
    End Sub
End Module"#, ["7"]);
vb_full_spec!(namespace_spec_imported_namespace_can_resolve_nested_class_constructor, r#"Namespace Demo.Core
    Public Class Box
        Public Value As String = "ctor"
    End Class
End Namespace
Imports Demo.Core
Module M
    Sub Main()
        Console.WriteLine((New Box()).Value)
    End Sub
End Module"#, ["ctor"]);
vb_full_spec!(namespace_spec_module_function_can_return_delegate, r#"Module Factory
    Public Function Build() As Func(Of Integer, Integer)
        Return Function(x) x + 1
    End Function
End Module
Module M
    Sub Main()
        Console.WriteLine(Factory.Build()(9))
    End Sub
End Module"#, ["10"]);
vb_full_spec!(namespace_spec_namespace_qualified_module_const_can_be_read, r#"Namespace Demo
    Public Module Config
        Public Const Name As String = "demo"
    End Module
End Namespace
Module M
    Sub Main()
        Console.WriteLine(Demo.Config.Name)
    End Sub
End Module"#, ["demo"]);
vb_full_spec!(namespace_spec_namespace_can_nest_three_levels_deep, r#"Namespace One
    Namespace Two
        Namespace Three
            Public Class Box
                Public Function Name() As String
                    Return "deep"
                End Function
            End Class
        End Namespace
    End Namespace
End Namespace
Module M
    Sub Main()
        Console.WriteLine((New One.Two.Three.Box()).Name())
    End Sub
End Module"#, ["deep"]);
