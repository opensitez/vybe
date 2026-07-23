use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Shared/Static Members Isolation Per Generic Type Instance
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_shared_field_per_type_isolation() {
    let src = r#"
Class TypeCounter(Of T)
    Public Shared Count As Integer = 0
End Class

Module Program
    Sub Main()
        TypeCounter(Of Integer).Count = 100
        TypeCounter(Of String).Count = 200
        Console.WriteLine(TypeCounter(Of Integer).Count & "|" & TypeCounter(Of String).Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|200"]);
}

#[test]
fn test_vb_generic_shared_property_per_type_isolation() {
    let src = r#"
Class PropertyHolder(Of T)
    Private Shared _data As String
    Public Shared Property Data As String
        Get
            Return _data
        End Get
        Set(value As String)
            _data = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        PropertyHolder(Of Integer).Data = "IntData"
        PropertyHolder(Of Double).Data = "DoubleData"
        Console.WriteLine(PropertyHolder(Of Integer).Data & "|" & PropertyHolder(Of Double).Data)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IntData|DoubleData"]);
}

#[test]
fn test_vb_generic_shared_method_per_type_access() {
    let src = r#"
Class TypeInfoProvider(Of T)
    Public Shared Function GetTypeName() As String
        Return GetType(T).Name
    End Function
End Class

Module Program
    Sub Main()
        Console.WriteLine(TypeInfoProvider(Of Integer).GetTypeName() & "|" & TypeInfoProvider(Of String).GetTypeName())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Int32|String"]);
}

#[test]
fn test_vb_generic_shared_event_per_type_isolation() {
    let src = r#"
Imports System

Class EventBus(Of T)
    Public Shared Event OnEvent As Action(Of T)
    Public Shared Sub Fire(item As T)
        RaiseEvent OnEvent(item)
    End Sub
End Class

Module Program
    Sub Main()
        AddHandler EventBus(Of Integer).OnEvent, Sub(i) Console.WriteLine("IntBus: " & i)
        AddHandler EventBus(Of String).OnEvent, Sub(s) Console.WriteLine("StringBus: " & s)

        EventBus(Of Integer).Fire(42)
        EventBus(Of String).Fire("Hello")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IntBus: 42", "StringBus: Hello"]);
}

#[test]
fn test_vb_generic_shared_constructor_per_type_counter() {
    let src = r#"
Class Tracker(Of T)
    Public Shared TotalInits As Integer = 0
    Shared Sub New()
        TotalInits += 1
    End Sub
End Class

Module Program
    Sub Main()
        ' Accessing Integer and Double triggers two distinct Shared Sub New invocations
        Console.WriteLine(Tracker(Of Integer).TotalInits & "|" & Tracker(Of Double).TotalInits)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|1"]);
}

#[test]
fn test_vb_generic_shared_list_cache_per_type() {
    let src = r#"
Imports System.Collections.Generic

Class Cache(Of T)
    Public Shared Items As New List(Of T)()
End Class

Module Program
    Sub Main()
        Cache(Of Integer).Items.Add(10)
        Cache(Of Integer).Items.Add(20)
        Cache(Of String).Items.Add("Alpha")

        Console.WriteLine(Cache(Of Integer).Items.Count & "|" & Cache(Of String).Items.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|1"]);
}

#[test]
fn test_vb_generic_shared_field_two_type_parameters() {
    let src = r#"
Class MatrixTracker(Of T1, T2)
    Public Shared InstanceID As String
End Class

Module Program
    Sub Main()
        MatrixTracker(Of Integer, String).InstanceID = "IntString"
        MatrixTracker(Of Integer, Double).InstanceID = "IntDouble"

        Console.WriteLine(MatrixTracker(Of Integer, String).InstanceID & "|" & MatrixTracker(Of Integer, Double).InstanceID)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IntString|IntDouble"]);
}

#[test]
fn test_vb_generic_shared_field_modified_via_instance() {
    let src = r#"
Class SharedAccess(Of T)
    Public Shared Tag As String = "Default"
End Class

Module Program
    Sub Main()
        Dim o1 As New SharedAccess(Of Integer)()
        Dim o2 As New SharedAccess(Of Integer)()
        SharedAccess(Of Integer).Tag = "Modified"

        Console.WriteLine(SharedAccess(Of Integer).Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Modified"]);
}

#[test]
fn test_vb_generic_shared_field_nested_class_access() {
    let src = r#"
Class Outer(Of T)
    Public Shared OuterData As String = "Outer"

    Public Class Inner
        Public Function GetOuterData() As String
            Return OuterData
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Outer(Of Integer).OuterData = "IntOuter"
        Outer(Of String).OuterData = "StringOuter"

        Dim inInt As New Outer(Of Integer).Inner()
        Dim inStr As New Outer(Of String).Inner()

        Console.WriteLine(inInt.GetOuterData() & "|" & inStr.GetOuterData())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IntOuter|StringOuter"]);
}

#[test]
fn test_vb_generic_shared_method_type_argument_inference() {
    let src = r#"
Module GenericHelper
    Public Function Identity(Of T)(item As T) As T
        Return item
    End Function
End Module

Module Program
    Sub Main()
        Console.WriteLine(GenericHelper.Identity(10) & "|" & GenericHelper.Identity("ABC"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|ABC"]);
}

#[test]
fn test_vb_generic_shared_method_static_local_variable() {
    let src = r#"
Module FunctionTracker
    Public Function Increment(Of T)() As Integer
        Static count As Integer = 0
        count += 1
        Return count
    End Function
End Module

Module Program
    Sub Main()
        Console.WriteLine(FunctionTracker.Increment(Of Integer)())
        Console.WriteLine(FunctionTracker.Increment(Of String)())
        Console.WriteLine(FunctionTracker.Increment(Of Integer)())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "2", "3"]);
}

#[test]
fn test_vb_generic_shared_field_custom_struct_value() {
    let src = r#"
Structure ConfigData
    Public ID As Integer
    Public Name As String
End Structure

Class SystemConfig(Of T)
    Public Shared Config As ConfigData
End Class

Module Program
    Sub Main()
        SystemConfig(Of Integer).Config = New ConfigData With {.ID = 1, .Name = "IntCfg"}
        SystemConfig(Of String).Config = New ConfigData With {.ID = 2, .Name = "StrCfg"}

        Console.WriteLine(SystemConfig(Of Integer).Config.Name & "|" & SystemConfig(Of String).Config.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IntCfg|StrCfg"]);
}

#[test]
fn test_vb_generic_shared_field_enum_isolation() {
    let src = r#"
Enum State
    Off = 0
    OnVal = 1
End Enum

Class StateHolder(Of T)
    Public Shared CurrentState As State = State.Off
End Class

Module Program
    Sub Main()
        StateHolder(Of Integer).CurrentState = State.OnVal
        Console.WriteLine(StateHolder(Of Integer).CurrentState.ToString() & "|" & StateHolder(Of String).CurrentState.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OnVal|Off"]);
}

#[test]
fn test_vb_generic_shared_dictionary_cache() {
    let src = r#"
Imports System.Collections.Generic

Class CacheRepository(Of TKey, TValue)
    Public Shared Lookup As New Dictionary(Of TKey, TValue)()
End Class

Module Program
    Sub Main()
        CacheRepository(Of String, Integer).Lookup("Key1") = 100
        CacheRepository(Of Integer, String).Lookup(1) = "One"

        Console.WriteLine(CacheRepository(Of String, Integer).Lookup("Key1") & "|" & CacheRepository(Of Integer, String).Lookup(1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100|One"]);
}

#[test]
fn test_vb_generic_shared_field_value_type_struct() {
    let src = r#"
Structure GenericStructHolder(Of T)
    Public Shared DefaultValue As T
End Structure

Module Program
    Sub Main()
        GenericStructHolder(Of Integer).DefaultValue = 99
        GenericStructHolder(Of String).DefaultValue = "DefaultText"

        Console.WriteLine(GenericStructHolder(Of Integer).DefaultValue & "|" & GenericStructHolder(Of String).DefaultValue)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99|DefaultText"]);
}

#[test]
fn test_vb_generic_shared_field_inheritance_base_class() {
    let src = r#"
Class BaseGeneric(Of T)
    Public Shared BaseTag As String = "Base"
End Class

Class DerivedInt
    Inherits BaseGeneric(Of Integer)
End Class

Class DerivedString
    Inherits BaseGeneric(Of String)
End Class

Module Program
    Sub Main()
        DerivedInt.BaseTag = "IntDerived"
        DerivedString.BaseTag = "StringDerived"

        Console.WriteLine(DerivedInt.BaseTag & "|" & DerivedString.BaseTag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IntDerived|StringDerived"]);
}

#[test]
fn test_vb_generic_shared_read_only_field_lazy_creation() {
    let src = r#"
Imports System

Class LazyContainer(Of T As New)
    Public Shared ReadOnly Instance As New T()
End Class

Class User
    Public Name As String = "DefaultUser"
End Class

Module Program
    Sub Main()
        Console.WriteLine(LazyContainer(Of User).Instance.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DefaultUser"]);
}

#[test]
fn test_vb_generic_shared_field_lock_object_isolation() {
    let src = r#"
Imports System

Class ThreadLockHolder(Of T)
    Public Shared ReadOnly SyncLockObject As New Object()
End Class

Module Program
    Sub Main()
        Dim lockInt = ThreadLockHolder(Of Integer).SyncLockObject
        Dim lockStr = ThreadLockHolder(Of String).SyncLockObject
        Console.WriteLine(Object.ReferenceEquals(lockInt, lockStr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_generic_shared_field_tuple_key() {
    let src = r#"
Class TupleStore(Of T)
    Public Shared StoredTuple As (String, T)
End Class

Module Program
    Sub Main()
        TupleStore(Of Integer).StoredTuple = ("Num", 42)
        TupleStore(Of String).StoredTuple = ("Text", "Val")

        Console.WriteLine(TupleStore(Of Integer).StoredTuple.Item2 & "|" & TupleStore(Of String).StoredTuple.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42|Val"]);
}

#[test]
fn test_vb_generic_shared_field_reflection_member_info() {
    let src = r#"
Class ReflectClass(Of T)
    Public Shared Counter As Integer = 0
End Class

Module Program
    Sub Main()
        Dim tInt = GetType(ReflectClass(Of Integer))
        Dim tStr = GetType(ReflectClass(Of String))
        Dim fieldInt = tInt.GetField("Counter")
        Dim fieldStr = tStr.GetField("Counter")

        fieldInt.SetValue(Nothing, 10)
        fieldStr.SetValue(Nothing, 20)

        Console.WriteLine(ReflectClass(Of Integer).Counter & "|" & ReflectClass(Of String).Counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|20"]);
}
