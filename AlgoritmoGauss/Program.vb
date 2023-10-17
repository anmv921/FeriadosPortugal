Imports System
Imports System.Security.Cryptography

Module Program
    Public Enum Meses
        Marco = 3
        Abril = 4
    End Enum

    Public Function MeeusJonesButcher(in_intAno As Integer) As DateTime
        Dim a, b, c, d, e, f, g, h, i, k, l, m As Integer
        Dim intDiaPascoa, intMesPascoa As Integer

        a = in_intAno Mod 19
        b = in_intAno \ 100
        c = in_intAno Mod 100
        d = b \ 4
        e = b Mod 4
        f = (b + 8) \ 25
        g = (b - f + 1) \ 3
        h = (19 * a + b - d - g + 15) Mod 30
        i = c \ 4
        k = c Mod 4
        l = (32 + 2 * e + 2 * i - h - k) Mod 7
        m = (a + 11 * h + 22 * l) \ 451

        intMesPascoa = (h + l - 7 * m + 114) \ 31
        intDiaPascoa = 1 + (h + l - 7 * m + 114) Mod 31

        Dim out_dtmDataPascoa As New DateTime(in_intAno, intMesPascoa, intDiaPascoa)
        Return out_dtmDataPascoa
    End Function

    Function AnoNovo(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=1, day:=1)
    End Function

    Function Entrudo(in_dtmPascoa As DateTime) As DateTime
        Return in_dtmPascoa.AddDays(-47)
    End Function

    Function SextaFeiraSanta(in_dtmPascoa) As DateTime
        Return in_dtmPascoa.AddDays(-2)
    End Function

    Function Pascoa(in_intAno As Integer) As DateTime
        Return MeeusJonesButcher(in_intAno)
    End Function

    Function DiaLiberdade(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=4, day:=25)
    End Function

    Function DiaTrabalhador(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=5, day:=1)
    End Function

    Function CorpoDeus(in_dtmPascoa As DateTime) As DateTime
        Return in_dtmPascoa.AddDays(60)
    End Function

    Function DiaPortugal(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=6, day:=10)
    End Function

    Function AssuncaoNossaSenhora(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=8, day:=15)
    End Function

    Function ImplantacaoRepublica(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=10, day:=5)
    End Function

    Function TodosOsSantos(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=11, day:=1)
    End Function

    Function RestauracaoIndependencia(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=12, day:=1)
    End Function

    Function ImaculadaConceicao(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=12, day:=8)
    End Function

    Function Natal(in_intAno As Integer) As DateTime
        Return New DateTime(year:=in_intAno, month:=12, day:=25)
    End Function

    Function ListaFeriados(in_intAno As Integer) As List(Of DateTime)
        Dim out_lstDtmFeriados As New List(Of DateTime)
        Dim dtmPascoa As DateTime = MeeusJonesButcher(in_intAno:=in_intAno)
        out_lstDtmFeriados.Add(AnoNovo(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(Entrudo(in_dtmPascoa:=dtmPascoa))
        out_lstDtmFeriados.Add(SextaFeiraSanta(in_dtmPascoa:=dtmPascoa))
        out_lstDtmFeriados.Add(dtmPascoa)
        out_lstDtmFeriados.Add(DiaLiberdade(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(DiaTrabalhador(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(CorpoDeus(in_dtmPascoa:=dtmPascoa))
        out_lstDtmFeriados.Add(DiaPortugal(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(AssuncaoNossaSenhora(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(ImplantacaoRepublica(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(TodosOsSantos(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(RestauracaoIndependencia(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(ImaculadaConceicao(in_intAno:=in_intAno))
        out_lstDtmFeriados.Add(Natal(in_intAno:=in_intAno))
        Return out_lstDtmFeriados
    End Function


    Sub Main(args As String())
        Dim strDate As String = "2024/05/01 08:30:52"
        Dim dtm As DateTime = DateTime.ParseExact(strDate, "yyyy/MM/dd HH:mm:ss",
                              System.Globalization.CultureInfo.InvariantCulture)
        Dim intAno As Integer = dtm.Year
        Dim dtmPascoa As DateTime
        Dim lst_dtmFeriados As New List(Of DateTime)
        dtmPascoa = MeeusJonesButcher(intAno)
        Console.WriteLine("Pascoa de " + intAno.ToString +
                          ": " + dtmPascoa.ToString("dd/MM/yyyy"))
        lst_dtmFeriados = ListaFeriados(in_intAno:=intAno)
        Console.WriteLine("Feriados do ano " + intAno.ToString + ":")
        For Each dtmFeriado As DateTime In lst_dtmFeriados
            Console.WriteLine(dtmFeriado.ToString("dd/MM/yyyy"))
        Next
    End Sub
End Module
