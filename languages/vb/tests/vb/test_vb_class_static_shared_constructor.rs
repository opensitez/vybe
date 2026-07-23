use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Shared Constructors (Shared Sub New) & Static Lifetime
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_shared_constructor_runs_once_on_first_access() {
    let src = r#"
Class StaticTracker
    Public Shared InitCount As Integer = 0
    Shared Sub New()
        InitCount += 1
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(StaticTracker.InitCount)
        Dim o1 As New StaticTracker()
        Dim o2 As New StaticTracker()
        Console.WriteLine(StaticTracker.InitCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "1"]);
}

#[test]
fn test_vb_shared_constructor_initializes_shared_fields() {
    let src = r#"
Class Config
    Public Shared ReadOnly StartTime As String
    Shared Sub New()
        StartTime = "2025-01-01 00:00:00"
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(Config.StartTime)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2025-01-01 00:00:00"]);
}

#[test]
fn test_vb_shared_constructor_in_generic_class_per_type_arg() {
    let src = r#"
Class GenericTracker(Of T)
    Public Shared Counter As Integer = 0
    Shared Sub New()
        Counter += 10
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(GenericTracker(Of Integer).Counter)
        Console.WriteLine(GenericTracker(Of String).Counter)
        Console.WriteLine(GenericTracker(Of Integer).Counter)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10", "10", "10"]);
}

#[test]
fn test_vb_shared_constructor_with_instance_constructor() {
    let src = r#"
Class Logger
    Public Shared Status As String
    Public Category As String

    Shared Sub New()
        Status = "SharedReady"
    End Sub

    Public Sub New(cat As String)
        Category = cat
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New Logger("File")
        Console.WriteLine(Logger.Status & "|" & l.Category)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SharedReady|File"]);
}

#[test]
fn test_vb_shared_constructor_exception_propagation() {
    let src = r#"
Imports System

Class FailingStatic
    Shared Sub New()
        Throw New InvalidOperationException("SharedInitFailed")
    End Sub
    Public Shared Sub Touch()
    End Sub
End Class

Module Program
    Sub Main()
        Try
            FailingStatic.Touch()
        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TypeInitializationException"]);
}

#[test]
fn test_vb_shared_constructor_struct_static_init() {
    let src = r#"
Structure Matrix
    Public Shared Identity As Matrix
    Public Data As String
    Shared Sub New()
        Identity = New Matrix With {.Data = "1,0,0,1"}
    End Sub
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Matrix.Identity.Data)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,0,0,1"]);
}

#[test]
fn test_vb_shared_constructor_base_derived_execution_order() {
    let src = r#"
Class ParentClass
    Public Shared Step1 As String
    Shared Sub New()
        Step1 = "ParentSharedInit"
        Console.WriteLine(Step1)
    End Sub
End Class

Class ChildClass
    Inherits ParentClass
    Public Shared Step2 As String
    Shared Sub New()
        Step2 = "ChildSharedInit"
        Console.WriteLine(Step2)
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New ChildClass()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ChildSharedInit", "ParentSharedInit"]);
}

#[test]
fn test_vb_shared_method_triggers_shared_constructor() {
    let src = r#"
Class Utility
    Public Shared Initialized As Boolean = False
    Shared Sub New()
        Initialized = True
    End Sub
    Public Shared Function Ping() As String
        Return "Pong"
    End Function
End Class

Module Program
    Sub Main()
        Dim res = Utility.Ping()
        Console.WriteLine(res & "|Initialized=" & Utility.Initialized)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Pong|Initialized=True"]);
}

#[test]
fn test_vb_shared_property_triggers_shared_constructor() {
    let src = r#"
Class SystemInfo
    Private Shared _version As String
    Shared Sub New()
        _version = "v1.0.0"
    End Sub
    Public Shared ReadOnly Property Version As String
        Get
            Return _version
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Console.WriteLine(SystemInfo.Version)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["v1.0.0"]);
}

#[test]
fn test_vb_shared_field_inline_initializer_vs_shared_constructor() {
    let src = r#"
Class OrderOfExecution
    Public Shared Value As Integer = 10
    Shared Sub New()
        Value = 20
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(OrderOfExecution.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_module_shared_constructor_equivalent() {
    let src = r#"
Module GlobalConfig
    Public AppName As String
    Sub New()
        AppName = "VybeApp"
    End Sub
End Module

Module Program
    Sub Main()
        Console.WriteLine(GlobalConfig.AppName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VybeApp"]);
}

#[test]
fn test_vb_shared_constructor_complex_collection_setup() {
    let src = r#"
Imports System.Collections.Generic

Class CacheManager
    Public Shared Lookup As New Dictionary(Of String, Integer)()
    Shared Sub New()
        Lookup.Add("Key1", 100)
        Lookup.Add("Key2", 200)
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(CacheManager.Lookup("Key1") & "+" & CacheManager.Lookup("Key2"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["300"]);
}

#[test]
fn test_vb_shared_constructor_private_shared_constructor() {
    let src = r#"
Class ConnectionPool
    Public Shared MaxConnections As Integer
    Private Shared Sub New()
        MaxConnections = 10
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(ConnectionPool.MaxConnections)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_shared_constructor_nested_class_trigger() {
    let src = r#"
Class Parent
    Public Shared ParentInit As Boolean = False
    Shared Sub New()
        ParentInit = True
    End Sub

    Public Class Child
        Public Shared Function Work() As String
            Return "ChildWork"
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim w = Parent.Child.Work()
        Console.WriteLine(w & "|ParentInit=" & Parent.ParentInit)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ChildWork|ParentInit=False"]);
}

#[test]
fn test_vb_shared_constructor_thread_safe_lazy_singleton() {
    let src = r#"
Class LazySingleton
    Public Shared ReadOnly Instance As LazySingleton
    Public Property CreatedAt As String
    Shared Sub New()
        Instance = New LazySingleton() With {.CreatedAt = "CreatedInSharedSubNew"}
    End Sub
    Private Sub New()
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(LazySingleton.Instance.CreatedAt)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CreatedInSharedSubNew"]);
}

#[test]
fn test_vb_shared_constructor_math_constants_precomputation() {
    let src = r#"
Imports System

Class MathLookup
    Public Shared ReadOnly SqrtTwo As Double
    Public Shared ReadOnly SqrtThree As Double
    Shared Sub New()
        SqrtTwo = Math.Sqrt(2.0)
        SqrtThree = Math.Sqrt(3.0)
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(Math.Round(MathLookup.SqrtTwo, 4) & "|" & Math.Round(MathLookup.SqrtThree, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.4142|1.7321"]);
}

#[test]
fn test_vb_shared_constructor_enum_string_dictionary_map() {
    let src = r#"
Imports System.Collections.Generic

Enum Priority
    Low
    High
End Enum

Class PriorityMapper
    Public Shared Map As New Dictionary(Of Priority, String)()
    Shared Sub New()
        Map(Priority.Low) = "Routine Priority"
        Map(Priority.High) = "Emergency Priority"
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(PriorityMapper.Map(Priority.High))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Emergency Priority"]);
}

#[test]
fn test_vb_shared_constructor_event_subscription() {
    let src = r#"
Imports System

Class SystemMonitor
    Public Shared Event Heartbeat As EventHandler
    Shared Sub New()
        ' Subscribe internal logger
        AddHandler Heartbeat, Sub(sender, args) Console.WriteLine("Internal Heartbeat Logged")
    End Sub
    Public Shared Sub Pulse()
        RaiseEvent Heartbeat(Nothing, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        SystemMonitor.Pulse()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Internal Heartbeat Logged"]);
}

#[test]
fn test_vb_shared_constructor_assembly_version_reader() {
    let src = r#"
Imports System.Reflection

Class VersionProvider
    Public Shared ReadOnly VersionString As String
    Shared Sub New()
        VersionString = GetType(VersionProvider).Assembly.GetName().Name
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(VersionProvider.VersionString IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_shared_constructor_reentry_safe() {
    let src = r#"
Class CircularA
    Public Shared Value As Integer
    Shared Sub New()
        Value = CircularB.Value + 10
    End Sub
End Class

Class CircularB
    Public Shared Value As Integer
    Shared Sub New()
        Value = 5
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine(CircularA.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["15"]);
}
