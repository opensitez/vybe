' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_payment_processor_strategy_pattern
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

Interface IPaymentStrategy
    Function ProcessPayment(amount As Decimal) As Boolean
End Interface

Class CreditCardPayment
    Implements IPaymentStrategy
    Public Function ProcessPayment(amount As Decimal) As Boolean Implements IPaymentStrategy.ProcessPayment
        Console.WriteLine("Paid $" & amount & " via CreditCard")
        Return True
    End Function
End Class

Class PayPalPayment
    Implements IPaymentStrategy
    Public Function ProcessPayment(amount As Decimal) As Boolean Implements IPaymentStrategy.ProcessPayment
        Console.WriteLine("Paid $" & amount & " via PayPal")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim strategy As IPaymentStrategy = New CreditCardPayment()
        strategy.ProcessPayment(45.5D)
    End Sub
End Module
