use super::helpers::run_vb;

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_full_spec!(object_model_spec_class_field_round_trips_value, r#"Class Person
    Public Name As String
End Class
Module M
    Sub Main()
        Dim p As New Person()
        p.Name = "Ada"
        Console.WriteLine(p.Name)
    End Sub
End Module"#, ["Ada"]);
vb_full_spec!(object_model_spec_constructor_sets_field_from_argument, r#"Class Counter
    Public Value As Integer
    Public Sub New(seed As Integer)
        Value = seed
    End Sub
End Class
Module M
    Sub Main()
        Console.WriteLine((New Counter(7)).Value)
    End Sub
End Module"#, ["7"]);
vb_full_spec!(object_model_spec_method_reads_instance_field, r#"Class Counter
    Public Value As Integer = 4
    Public Function Current() As Integer
        Return Value
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New Counter()).Current())
    End Sub
End Module"#, ["4"]);
vb_full_spec!(object_model_spec_method_updates_instance_field, r#"Class Counter
    Public Value As Integer
    Public Sub Increment()
        Value += 1
    End Sub
End Class
Module M
    Sub Main()
        Dim c As New Counter()
        c.Increment()
        c.Increment()
        Console.WriteLine(c.Value)
    End Sub
End Module"#, ["2"]);
vb_full_spec!(object_model_spec_overloaded_constructors_choose_parameterized_version, r#"Class Box
    Public Text As String
    Public Sub New()
        Text = "empty"
    End Sub
    Public Sub New(value As String)
        Text = value
    End Sub
End Class
Module M
    Sub Main()
        Console.WriteLine((New Box("filled")).Text)
    End Sub
End Module"#, ["filled"]);
vb_full_spec!(object_model_spec_overloaded_methods_dispatch_by_arity, r#"Class Printer
    Public Function Show() As String
        Return "zero"
    End Function
    Public Function Show(value As Integer) As String
        Return CStr(value)
    End Function
End Class
Module M
    Sub Main()
        Dim p As New Printer()
        Console.WriteLine(p.Show())
        Console.WriteLine(p.Show(9))
    End Sub
End Module"#, ["zero", "9"]);
vb_full_spec!(object_model_spec_overloaded_methods_dispatch_by_parameter_type, r#"Class Printer
    Public Function Show(value As Integer) As String
        Return "int:" & value
    End Function
    Public Function Show(value As String) As String
        Return "str:" & value
    End Function
End Class
Module M
    Sub Main()
        Dim p As New Printer()
        Console.WriteLine(p.Show(5))
        Console.WriteLine(p.Show("vb"))
    End Sub
End Module"#, ["int:5", "str:vb"]);
vb_full_spec!(object_model_spec_auto_property_round_trips_value, r#"Class Person
    Public Property Name As String
End Class
Module M
    Sub Main()
        Dim p As New Person()
        p.Name = "Grace"
        Console.WriteLine(p.Name)
    End Sub
End Module"#, ["Grace"]);
vb_full_spec!(object_model_spec_read_only_property_returns_backing_field, r#"Class Bag
    Private _count As Integer = 3
    Public ReadOnly Property Count As Integer
        Get
            Return _count
        End Get
    End Property
End Class
Module M
    Sub Main()
        Console.WriteLine((New Bag()).Count)
    End Sub
End Module"#, ["3"]);
vb_full_spec!(object_model_spec_write_only_property_updates_backing_field, r#"Class Bag
    Private _count As Integer
    Public WriteOnly Property Count As Integer
        Set(value As Integer)
            _count = value
        End Set
    End Property
    Public Function Snapshot() As Integer
        Return _count
    End Function
End Class
Module M
    Sub Main()
        Dim b As New Bag()
        b.Count = 5
        Console.WriteLine(b.Snapshot())
    End Sub
End Module"#, ["5"]);
vb_full_spec!(object_model_spec_custom_property_getter_transforms_value, r#"Class Person
    Private _name As String = "ada"
    Public ReadOnly Property UpperName As String
        Get
            Return UCase(_name)
        End Get
    End Property
End Class
Module M
    Sub Main()
        Console.WriteLine((New Person()).UpperName)
    End Sub
End Module"#, ["ADA"]);
vb_full_spec!(object_model_spec_custom_property_setter_transforms_value, r#"Class Person
    Private _name As String
    Public WriteOnly Property LowerName As String
        Set(value As String)
            _name = LCase(value)
        End Set
    End Property
    Public Function Snapshot() As String
        Return _name
    End Function
End Class
Module M
    Sub Main()
        Dim p As New Person()
        p.LowerName = "Ada"
        Console.WriteLine(p.Snapshot())
    End Sub
End Module"#, ["ada"]);
vb_full_spec!(object_model_spec_default_property_can_index_collection_like_type, r#"Class Buffer
    Private values() As String = {"a", "b", "c"}
    Default Public Property Item(index As Integer) As String
        Get
            Return values(index)
        End Get
        Set(value As String)
            values(index) = value
        End Set
    End Property
End Class
Module M
    Sub Main()
        Dim buffer As New Buffer()
        buffer(1) = "x"
        Console.WriteLine(buffer(1))
    End Sub
End Module"#, ["x"]);
vb_full_spec!(object_model_spec_structure_field_round_trips_value, r#"Structure Point
    Public X As Integer
End Structure
Module M
    Sub Main()
        Dim p As Point
        p.X = 7
        Console.WriteLine(p.X)
    End Sub
End Module"#, ["7"]);
vb_full_spec!(object_model_spec_structure_method_reads_state, r#"Structure Point
    Public X As Integer
    Public Function DoubleX() As Integer
        Return X * 2
    End Function
End Structure
Module M
    Sub Main()
        Dim p As Point
        p.X = 6
        Console.WriteLine(p.DoubleX())
    End Sub
End Module"#, ["12"]);
vb_full_spec!(object_model_spec_structure_passed_by_value_keeps_original, r#"Structure Point
    Public X As Integer
End Structure
Module M
    Sub MoveRight(p As Point)
        p.X += 1
    End Sub
    Sub Main()
        Dim p As Point
        p.X = 3
        MoveRight(p)
        Console.WriteLine(p.X)
    End Sub
End Module"#, ["3"]);
vb_full_spec!(object_model_spec_structure_passed_byref_can_mutate_original, r#"Structure Point
    Public X As Integer
End Structure
Module M
    Sub MoveRight(ByRef p As Point)
        p.X += 1
    End Sub
    Sub Main()
        Dim p As Point
        p.X = 3
        MoveRight(p)
        Console.WriteLine(p.X)
    End Sub
End Module"#, ["4"]);
vb_full_spec!(object_model_spec_enum_values_compare_equal_to_named_member, r#"Enum Color
    Red = 1
    Blue = 2
End Enum
Module M
    Sub Main()
        Console.WriteLine(Color.Red = 1)
    End Sub
End Module"#, ["true"]);
vb_full_spec!(object_model_spec_enum_underlying_value_prints_numeric_representation, r#"Enum Color
    Red = 4
End Enum
Module M
    Sub Main()
        Console.WriteLine(Color.Red)
    End Sub
End Module"#, ["4"]);
vb_full_spec!(object_model_spec_interface_method_dispatch_uses_implementation, r#"Interface IGreeter
    Function Greet() As String
End Interface
Class Greeter
    Implements IGreeter
    Public Function Greet() As String Implements IGreeter.Greet
        Return "hello"
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New Greeter()).Greet())
    End Sub
End Module"#, ["hello"]);
vb_full_spec!(object_model_spec_interface_reference_can_call_method, r#"Interface IGreeter
    Function Greet() As String
End Interface
Class Greeter
    Implements IGreeter
    Public Function Greet() As String Implements IGreeter.Greet
        Return "hello"
    End Function
End Class
Module M
    Sub Main()
        Dim value As IGreeter = New Greeter()
        Console.WriteLine(value.Greet())
    End Sub
End Module"#, ["hello"]);
vb_full_spec!(object_model_spec_mustinherit_base_can_be_used_through_derived_type, r#"MustInherit Class Shape
    Public MustOverride Function Name() As String
End Class
Class Circle
    Inherits Shape
    Public Overrides Function Name() As String
        Return "circle"
    End Function
End Class
Module M
    Sub Main()
        Dim value As Shape = New Circle()
        Console.WriteLine(value.Name())
    End Sub
End Module"#, ["circle"]);
vb_full_spec!(object_model_spec_overridable_method_can_be_overridden, r#"Class BasePrinter
    Public Overridable Function Speak() As String
        Return "base"
    End Function
End Class
Class LoudPrinter
    Inherits BasePrinter
    Public Overrides Function Speak() As String
        Return "loud"
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New LoudPrinter()).Speak())
    End Sub
End Module"#, ["loud"]);
vb_full_spec!(object_model_spec_mustoverride_method_is_implemented_by_derived_class, r#"MustInherit Class Shape
    Public MustOverride Function Area() As Integer
End Class
Class Square
    Inherits Shape
    Public Overrides Function Area() As Integer
        Return 9
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New Square()).Area())
    End Sub
End Module"#, ["9"]);
vb_full_spec!(object_model_spec_mybase_calls_base_implementation, r#"Class BasePrinter
    Public Overridable Function Speak() As String
        Return "base"
    End Function
End Class
Class LoudPrinter
    Inherits BasePrinter
    Public Overrides Function Speak() As String
        Return MyBase.Speak() & "+derived"
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New LoudPrinter()).Speak())
    End Sub
End Module"#, ["base+derived"]);
vb_full_spec!(object_model_spec_me_reference_returns_current_instance_state, r#"Class Counter
    Public Value As Integer = 5
    Public Function Snapshot() As Integer
        Return Me.Value
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New Counter()).Snapshot())
    End Sub
End Module"#, ["5"]);
vb_full_spec!(object_model_spec_property_can_return_object_field, r#"Class Address
    Public City As String
End Class
Class Person
    Public Home As Address
End Class
Module M
    Sub Main()
        Dim p As New Person()
        p.Home = New Address()
        p.Home.City = "Paris"
        Console.WriteLine(p.Home.City)
    End Sub
End Module"#, ["Paris"]);
vb_full_spec!(object_model_spec_object_initializer_sets_field, r#"Class Person
    Public Name As String
End Class
Module M
    Sub Main()
        Dim p As New Person() With {.Name = "Ada"}
        Console.WriteLine(p.Name)
    End Sub
End Module"#, ["Ada"]);
vb_full_spec!(object_model_spec_object_initializer_sets_property, r#"Class Person
    Public Property Name As String
End Class
Module M
    Sub Main()
        Dim p As New Person() With {.Name = "Grace"}
        Console.WriteLine(p.Name)
    End Sub
End Module"#, ["Grace"]);
vb_full_spec!(object_model_spec_nested_class_can_be_instantiated_through_outer_type, r#"Class Outer
    Public Class Inner
        Public Function Name() As String
            Return "inner"
        End Function
    End Class
End Class
Module M
    Sub Main()
        Console.WriteLine((New Outer.Inner()).Name())
    End Sub
End Module"#, ["inner"]);
vb_full_spec!(object_model_spec_generic_class_stores_integer_value, r#"Class Box(Of T)
    Public Value As T
    Public Sub New(v As T)
        Value = v
    End Sub
End Class
Module M
    Sub Main()
        Console.WriteLine((New Box(Of Integer)(7)).Value)
    End Sub
End Module"#, ["7"]);
vb_full_spec!(object_model_spec_generic_class_stores_string_value, r#"Class Box(Of T)
    Public Value As T
    Public Sub New(v As T)
        Value = v
    End Sub
End Class
Module M
    Sub Main()
        Console.WriteLine((New Box(Of String)("vb")).Value)
    End Sub
End Module"#, ["vb"]);
vb_full_spec!(object_model_spec_generic_method_echoes_argument, r#"Class Echoer
    Public Function Echo(Of T)(value As T) As T
        Return value
    End Function
End Class
Module M
    Sub Main()
        Console.WriteLine((New Echoer()).Echo(Of String)("hello"))
    End Sub
End Module"#, ["hello"]);
vb_full_spec!(object_model_spec_class_can_contain_list_field, r#"Class Bag
    Public Items As New List(Of Integer)
End Class
Module M
    Sub Main()
        Dim bag As New Bag()
        bag.Items.Add(9)
        Console.WriteLine(bag.Items.Count)
    End Sub
End Module"#, ["1"]);
vb_full_spec!(object_model_spec_constructor_can_allocate_list_field, r#"Class Bag
    Public Items As List(Of Integer)
    Public Sub New()
        Items = New List(Of Integer)()
    End Sub
End Class
Module M
    Sub Main()
        Dim bag As New Bag()
        bag.Items.Add(4)
        Console.WriteLine(bag.Items(0))
    End Sub
End Module"#, ["4"]);
vb_full_spec!(object_model_spec_method_can_return_me_for_fluent_usage, r#"Class Counter
    Public Value As Integer
    Public Function Increment() As Counter
        Value += 1
        Return Me
    End Function
End Class
Module M
    Sub Main()
        Dim c As New Counter()
        Console.WriteLine(c.Increment().Increment().Value)
    End Sub
End Module"#, ["2"]);
vb_full_spec!(object_model_spec_nullable_integer_property_can_hold_nothing, r#"Class Holder
    Public Property Value As Integer?
End Class
Module M
    Sub Main()
        Dim h As New Holder()
        h.Value = Nothing
        Console.WriteLine(IsNothing(h.Value))
    End Sub
End Module"#, ["true"]);
vb_full_spec!(object_model_spec_class_can_expose_array_property, r#"Class Holder
    Public Property Values As Integer()
End Class
Module M
    Sub Main()
        Dim h As New Holder()
        h.Values = New Integer() {1, 2, 3}
        Console.WriteLine(h.Values(2))
    End Sub
End Module"#, ["3"]);
vb_full_spec!(object_model_spec_property_can_compute_from_two_fields, r#"Class Pair
    Public LeftValue As Integer
    Public RightValue As Integer
    Public ReadOnly Property Sum As Integer
        Get
            Return LeftValue + RightValue
        End Get
    End Property
End Class
Module M
    Sub Main()
        Dim pair As New Pair()
        pair.LeftValue = 3
        pair.RightValue = 4
        Console.WriteLine(pair.Sum)
    End Sub
End Module"#, ["7"]);
vb_full_spec!(object_model_spec_private_backing_field_is_hidden_behind_property, r#"Class Meter
    Private _value As Integer
    Public Property Value As Integer
        Get
            Return _value
        End Get
        Set(value As Integer)
            _value = value
        End Set
    End Property
End Class
Module M
    Sub Main()
        Dim meter As New Meter()
        meter.Value = 12
        Console.WriteLine(meter.Value)
    End Sub
End Module"#, ["12"]);
vb_full_spec!(object_model_spec_base_reference_can_point_to_derived_instance, r#"Class Animal
    Public Overridable Function Speak() As String
        Return "animal"
    End Function
End Class
Class Dog
    Inherits Animal
    Public Overrides Function Speak() As String
        Return "dog"
    End Function
End Class
Module M
    Sub Main()
        Dim pet As Animal = New Dog()
        Console.WriteLine(pet.Speak())
    End Sub
End Module"#, ["dog"]);
vb_full_spec!(object_model_spec_interface_can_be_implemented_by_structure, r#"Interface INameable
    Function Name() As String
End Interface
Structure Thing
    Implements INameable
    Public Function Name() As String Implements INameable.Name
        Return "thing"
    End Function
End Structure
Module M
    Sub Main()
        Dim value As INameable = New Thing()
        Console.WriteLine(value.Name())
    End Sub
End Module"#, ["thing"]);
vb_full_spec!(object_model_spec_structure_property_round_trips_value, r#"Structure Meter
    Public Property Value As Integer
End Structure
Module M
    Sub Main()
        Dim meter As Meter
        meter.Value = 8
        Console.WriteLine(meter.Value)
    End Sub
End Module"#, ["8"]);
vb_full_spec!(object_model_spec_enum_can_drive_select_case_branch, r#"Enum Tone
    Low
    High
End Enum
Module M
    Sub Main()
        Dim tone As Tone = Tone.High
        Select Case tone
            Case Tone.Low
                Console.WriteLine("low")
            Case Tone.High
                Console.WriteLine("high")
        End Select
    End Sub
End Module"#, ["high"]);
vb_full_spec!(object_model_spec_class_can_contain_const_field, r#"Class Config
    Public Const AppName As String = "Vybe"
End Class
Module M
    Sub Main()
        Console.WriteLine(Config.AppName)
    End Sub
End Module"#, ["Vybe"]);
vb_full_spec!(object_model_spec_notinheritable_class_can_be_instantiated_normally, r#"NotInheritable Class Token
    Public Value As String = "sealed"
End Class
Module M
    Sub Main()
        Console.WriteLine((New Token()).Value)
    End Sub
End Module"#, ["sealed"]);
vb_full_spec!(object_model_spec_shadows_member_can_hide_base_field, r#"Class BaseToken
    Public Name As String = "base"
End Class
Class DerivedToken
    Inherits BaseToken
    Public Shadows Name As String = "derived"
End Class
Module M
    Sub Main()
        Console.WriteLine((New DerivedToken()).Name)
    End Sub
End Module"#, ["derived"]);
vb_full_spec!(object_model_spec_property_setter_can_ignore_redundant_assignment, r#"Class Meter
    Private _value As Integer
    Public Property Value As Integer
        Get
            Return _value
        End Get
        Set(value As Integer)
            If _value <> value Then
                _value = value
            End If
        End Set
    End Property
End Class
Module M
    Sub Main()
        Dim meter As New Meter()
        meter.Value = 3
        meter.Value = 3
        Console.WriteLine(meter.Value)
    End Sub
End Module"#, ["3"]);
vb_full_spec!(object_model_spec_object_collection_property_can_be_iterated, r#"Class Bag
    Public Property Items As List(Of String)
End Class
Module M
    Sub Main()
        Dim bag As New Bag()
        bag.Items = New List(Of String)()
        bag.Items.Add("a")
        bag.Items.Add("b")
        Dim text As String = ""
        For Each item In bag.Items
            text &= item
        Next
        Console.WriteLine(text)
    End Sub
End Module"#, ["ab"]);
vb_full_spec!(object_model_spec_property_can_store_structure_value, r#"Structure Point
    Public X As Integer
End Structure
Class Holder
    Public Property Location As Point
End Class
Module M
    Sub Main()
        Dim holder As New Holder()
        holder.Location = New Point With {.X = 9}
        Console.WriteLine(holder.Location.X)
    End Sub
End Module"#, ["9"]);
