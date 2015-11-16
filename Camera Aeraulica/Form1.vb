Option Explicit On
'Option Strict On

Imports System.Net
Imports System.IO
Imports System.Net.Sockets
Imports System.Text
Imports System.IO.Ports
Imports System.Math
Imports Camera_Aeraulica.strumenti
Imports Microsoft.Office.Interop.Excel



Public Class Form1

    '***************************************************************************************************************
    '----------------------------- INIZIO DICHIARAZIONI VARIABILI GLOBALI ------------------------------------------
    '***************************************************************************************************************
    '20151115   versione iniziale caricata su Github

    'serve per verificare la durata dei processi
    Public sw As New Stopwatch

    'variabili per il report Excel

    Public xlApp As Microsoft.Office.Interop.Excel.Application
    Public xlBook As Microsoft.Office.Interop.Excel.Workbook
    Public xlSheet As Microsoft.Office.Interop.Excel.Worksheet

    'variabili relative agli strumenti
    Public stazione_barometrica As New c_stazione_barometrica
    Public tachimetro As New c_tachimetro
    Public aux_fan As New c_analogout
    Public trasm_P1 As New c_trasmettitore_pressione
    Public trasm_DP As New c_trasmettitore_pressione
    Public wattmetro As New c_wattmetro
    Public esam As New c_esam_e2002
    Public multimetro As New c_multimetro
    Public Ok_Analog, Ok_Barometro, Ok_Tacho, Ok_P1, Ok_DP, Ok_Wattmetro, Ok_Multimetro As Boolean
    Public IPBarometro As String
    Public nome_porta_P1, nome_porta_DP, nome_porta_tachimetro, nome_porta_wattmetro, nome_porta_multimetro, nome_porta_esam As String

    Public Const Pmax_assoluta As Double = 500 'fondoscala P1
    Public Vaux_percent_max As Double = 100
    Public Const Vaux_percent_min As Double = 8

    Public Const serranda_max As Double = 90 'in realtà non è il percento ma sono i gradi
    Public Const serranda_min As Double = 0

    'variabili relative alla configurazione ugelli
    Structure ugello
        Dim diametro As Single
        Dim aperto As Boolean
        Dim Q_max As Integer 'portata massima misurabile con quell'ugello e DP=490Pa
        Dim L_su_d As Single
        Dim k1_alfa As Double 'dipende da L/d - vedere formula per il calcolo di alfa - pg.70 ISO 5801
        Dim k2_alfa As Double 'dipende da L/d - vedere formula per il calcolo di alfa - pg.70 ISO 5801
    End Structure
    Public Const N_ugelli As Byte = 6
    Public Const N_configurazioni As Byte = (2 ^ N_ugelli)  'le configurazioni teoricamente possibili con N_ugelli appena impostato
    Dim ugelli(N_ugelli - 1) As ugello 'in VB gli array si dichiarano array(dimensione-1)
    Public Conf_Ug As UInt16 'il numero binario corrispondente alla configurazione ugelli
    Dim configurazioni_significative(N_configurazioni - 1) As UInt16 'un array che contiene le configurazioni proposte dal programma
    Public immagine_ugello_aperto = Camera_Aeraulica.My.Resources.Resources.ugello_aperto
    Public immagine_ugello_chiuso = Camera_Aeraulica.My.Resources.Resources.ugello_chiuso

    'variabili di configurazione prova
    Public Qmax_assoluta As UInt32 'portata massima assoluta misurabile dalla camera
    Public Qmax, Larghezza, Altezza As Double
    Public Pa_calc As Double = 101325
    Public Hu_calc_percent As Double = 40
    Public Ta_calc As Double = 20
    Public Roref As Double = 1.2
    Public Install, Data, Path, NomeFile As String
    Public num_totale_punti As Integer

    'parametri controllo e sistema
    Public Const Tcampionamento As Double = 1
    Public Tc_portata, Tc_pressione, Kp_portata, Kp_pressione, Kc_portata, Kc_pressione, Kconfig_portata, Kconfig_pressione As Double
    Public Tp, Tetap As Double
    Public Const Kconfig_pressione_01 As Double = -1000

    'PARAMETRI DI CORREZIONE INTRODOTTI DOPO LA PROVA COMPARATIVA CON TUV
    'LA13016
    Public Const coeff_TUV_portata As Single = 1.03
    Public Const coeff_TUV_pressione As Single = 1.05

    'variabili relative all'acquisizione dati e ai calcoli
    Public ciclo_lettura As Byte = 1
    Public ratio_raggiungimento_punto As Double
    Public Dp, n, We_p, I_p, V_p, Pa, Ta, Hu, P_ing_ug, P_cam1, P_cam2, T_ing_ug, T_cam2 As Double
    Public P6, P4, Teta6, Ro6, Psg3, P7, Teta7, Ro7, Teta3, Ro3, A2, Psat_Ta, Pv, Teta_a, Roa, Rw, Ro1, qm, Sum_Cj_d2, somma_d2_ugelli As Double
    Public qv, qv_max, Ps, Pd, Pt, We, Wa, ni, I, V As Double
    Public qv_pr, Ps_pr, Pd_pr, Pt_pr, We_pr, Wa_pr, ni_pr, I_pr, V_pr As Double
    Public k As Double = 1.4
    Public epsilon As Double = 1
    Public beta As Double = 1
    Public Err_prev, Err_act, DQ As Double
    Public Fase, Press_Target_Ok, Q_Target_Ok, Vaux_target_Ok, nq, StandBy, Limit, TestMaxPress As Integer
    Public Qman, Pman, Vaux_percent, serranda_percent As Double
    Public Test_Manuale As Boolean
    Public Stop_Max_Press As Boolean = 0
    Public pressioni_azzerate As Boolean = False

    Public Ptemp As Double
    Public Ptemp_prev As Double = 0
    Public qm_prev As Double = 0
    Public dp_prev As Double = 0
    Public V_p_prev As Double = 0
    Public We_p_prev As Double = 0
    Public I_p_prev As Double = 0
    Public Vaux_percent_prev As Double = 0
    Public serranda_percent_prev As Double = 0
    Public Status_prev, Path_prev As String

    Public Const num_max_punti As Integer = 20
    Public Const num_cicli_lettura_stabile As Integer = 20  'faccio una media di 20s prima di acquisire (da 0 a 19)
    Public Const num_max_punti_portata As Integer = 18

    Public array_20_vuoto() As Double = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
    Public array_3_vuoto() As Double = {0, 0, 0}
    Public array_punti_portata() As Double = {}

    'array per media valori
    Public step_buf_P_cam1, step_buf_P_cam2, step_buf_Dp, step_buf_TACHO, step_buf_V_p, step_buf_I_p, step_buf_We_p As Integer
    Public Buffer_Dp() As Double = array_3_vuoto
    Public Buffer_P_cam1() As Double = array_3_vuoto
    Public Buffer_P_cam2() As Double = array_3_vuoto
    Public Buffer_TACHO() As Double = array_3_vuoto
    Public Buffer_V_p() As Double = array_3_vuoto
    Public Buffer_I_p() As Double = array_3_vuoto
    Public Buffer_We_p() As Double = array_3_vuoto

    'array punti acquisiti - dati corretti
    Public Ps_Array() As Double = array_20_vuoto
    Public Qv_Array() As Double = array_20_vuoto
    Public Ro_Array() As Double = array_20_vuoto
    Public Pd_Array() As Double = array_20_vuoto
    Public Pt_Array() As Double = array_20_vuoto
    Public n_Array() As Double = array_20_vuoto
    Public I_Array() As Double = array_20_vuoto
    Public We_Array() As Double = array_20_vuoto
    Public Wa_Array() As Double = array_20_vuoto
    Public EFF_Array() As Double = array_20_vuoto
    Public V_Array() As Double = array_20_vuoto

    'array punti acquisiti - dati alle condizioni di prova
    Public Ps_pr_Array() As Double = array_20_vuoto
    Public Qv_pr_Array() As Double = array_20_vuoto
    Public Ro_pr_Array() As Double = array_20_vuoto
    Public Pd_pr_Array() As Double = array_20_vuoto
    Public Pt_pr_Array() As Double = array_20_vuoto
    Public We_pr_Array() As Double = array_20_vuoto
    Public Wa_pr_Array() As Double = array_20_vuoto
    Public EFF_pr_Array() As Double = array_20_vuoto
    Public I_pr_Array() As Double = array_20_vuoto
    Public V_pr_Array() As Double = array_20_vuoto

    'buffer dati acquisiti per media prima di salvataggio dato in report
    Public Ta_buffer_circolare As New List(Of Double)
    Public T_ing_ug_buffer_circolare As New List(Of Double)
    Public T_cam2_buffer_circolare As New List(Of Double)
    Public Hu_buffer_circolare As New List(Of Double)
    Public Pa_buffer_circolare As New List(Of Double)
    Public V_p_buffer_circolare As New List(Of Double)
    Public I_p_buffer_circolare As New List(Of Double)
    Public We_p_buffer_circolare As New List(Of Double)
    Public n_buffer_circolare As New List(Of Double)
    Public Dp_buffer_circolare As New List(Of Double)
    Public P_cam1_buffer_circolare As New List(Of Double)
    Public P_cam2_buffer_circolare As New List(Of Double)
    Public P_ing_ug_buffer_circolare As New List(Of Double)

    '***************************************************************************************************************
    '----------------------------- FINE DICHIARAZIONI VARIABILI GLOBALI --------------------------------------------
    '***************************************************************************************************************

    'MEDIA DATI SU TRE ACQUISIZIONI
    'OK
    Function Filtro_Dati(ByVal Data_read As Double, ByRef Data_read_buf() As Double, ByRef step_buf As Integer) As Double

        Data_read_buf(step_buf) = Data_read
        step_buf = step_buf + 1
        If step_buf > 2 Then step_buf = 0

        Return (Data_read_buf(0) + Data_read_buf(1) + Data_read_buf(2)) / 3 'media su 3 valori

    End Function

    'ARROTONDAMENTO DECIMALI
    'OK
    Sub Virgola(ByRef valore As Double, ByVal decimali As Integer)

        valore = Round(valore, decimali)

    End Sub
    'PROCEDURE RELATIVE AI BUFFER DI LETTURA

    Private Sub shifta_indietro_buffer_circolari()

        For Each Buffer As List(Of Double) In {Ta_buffer_circolare, _
                                        T_ing_ug_buffer_circolare, _
                                        T_cam2_buffer_circolare, _
                                        Hu_buffer_circolare, _
                                        Pa_buffer_circolare, _
                                        V_p_buffer_circolare, _
                                        I_p_buffer_circolare, _
                                        We_p_buffer_circolare, _
                                        n_buffer_circolare, _
                                        Dp_buffer_circolare, _
                                        P_cam1_buffer_circolare, _
                                        P_cam2_buffer_circolare, _
                                        P_ing_ug_buffer_circolare}

            'per ogni lista elimino l'elemento più vecchio
            If Buffer.Count > (num_cicli_lettura_stabile - 1) Then Buffer.RemoveAt(0)
            Buffer.TrimExcess()
        Next
    End Sub

    Private Sub reset_buffer_circolari()

        Ta_buffer_circolare.Clear()
        T_ing_ug_buffer_circolare.Clear()
        T_cam2_buffer_circolare.Clear()
        Hu_buffer_circolare.Clear()
        Pa_buffer_circolare.Clear()
        V_p_buffer_circolare.Clear()
        I_p_buffer_circolare.Clear()
        We_p_buffer_circolare.Clear()
        n_buffer_circolare.Clear()
        Dp_buffer_circolare.Clear()
        P_cam1_buffer_circolare.Clear()
        P_cam2_buffer_circolare.Clear()
        P_ing_ug_buffer_circolare.Clear()

    End Sub

    Private Sub riempi_buffer_circolari(ByVal indice As Integer)

        If indice > num_cicli_lettura_stabile Then 'se sono già oltre la stabilità continuo ad acquisire eliminando la lettura più vecchia
            shifta_indietro_buffer_circolari()
        End If

        If indice > 0 Then
            'immetto le letture nei buffer circolari
            Ta_buffer_circolare.Add(Ta)
            T_ing_ug_buffer_circolare.Add(T_ing_ug)
            T_cam2_buffer_circolare.Add(T_cam2)
            Hu_buffer_circolare.Add(Hu)
            Pa_buffer_circolare.Add(Pa)
            V_p_buffer_circolare.Add(V_p)
            I_p_buffer_circolare.Add(I_p)
            We_p_buffer_circolare.Add(We_p)
            n_buffer_circolare.Add(n)
            Dp_buffer_circolare.Add(Dp)
            P_cam1_buffer_circolare.Add(P_cam1)
            P_cam2_buffer_circolare.Add(P_cam2)
            P_ing_ug_buffer_circolare.Add(P_ing_ug)

        Else
            If (Ta_buffer_circolare.Count > 0) Then reset_buffer_circolari()
            'se l'indice è 0 e c'è qualcosa nei buffer viene scartato
        End If


    End Sub


    Private Sub media_buffer_circolari()

        'metto nelle variabili la media dei buffer circolari in modo da salvare dati più stabilizzati

        Ta = Ta_buffer_circolare.Average
        T_ing_ug = T_ing_ug_buffer_circolare.Average
        T_cam2 = T_cam2_buffer_circolare.Average
        Hu = Hu_buffer_circolare.Average
        Pa = Pa_buffer_circolare.Average
        V_p = V_p_buffer_circolare.Average
        Virgola(V_p, 1)
        I_p = I_p_buffer_circolare.Average
        Virgola(I_p, 3)
        We_p = We_p_buffer_circolare.Average
        Virgola(We_p, 1)
        n = n_buffer_circolare.Average
        Virgola(n, 0)
        Dp = Dp_buffer_circolare.Average
        Virgola(Dp, 1)
        P_cam1 = P_cam1_buffer_circolare.Average
        Virgola(P_cam1, 1)
        P_cam2 = P_cam2_buffer_circolare.Average
        Virgola(P_cam2, 1)
        P_ing_ug = P_ing_ug_buffer_circolare.Average

        'cancello i buffer
        reset_buffer_circolari()

    End Sub


    'RIEMPIMENTO PAGINA DI CONFIGURAZIONE DELLE PORTE SERIALI PER I DIVERSI SENSORI COLLEGATI
    'OK
    Public Sub CaricaImpostazioni(ByVal stringa_impostazioni As String)

        Dim Dati() As String = Split(stringa_impostazioni, Chr(9))

        tb_IP1.Text = Dati(0)
        tb_IP2.Text = Dati(1)
        tb_IP3.Text = Dati(2)
        tb_IP4.Text = Dati(3)
        cb_porta_P1.SelectedItem = Dati(4)
        cb_porta_DP.SelectedItem = Dati(5)
        cb_porta_wattmetro.SelectedItem = Dati(6)
        cb_porta_tachimetro.SelectedItem = Dati(7)
        cb_porta_multimetro.SelectedItem = Dati(8)

        tb_percorso_file_report.Text = Dati(9)
        Path = Dati(9)

    End Sub

    'RIEMPIMENTIO ARRAY DATI SU ACQUISIZIONE DI UN PUNTO DURANTE LA PROVA
    'OK
    Public Sub PopolaArray(ByVal a As Integer)

        Ro_Array(a) = Roref
        Qv_Array(a) = qv
        Ps_Array(a) = Ps
        Pd_Array(a) = Pd
        Pt_Array(a) = Pt
        n_Array(a) = n
        I_Array(a) = I
        We_Array(a) = We
        Wa_Array(a) = Wa
        EFF_Array(a) = ni
        V_Array(a) = V

        Ro_pr_Array(a) = Roa
        Qv_pr_Array(a) = qv_pr
        Ps_pr_Array(a) = Ps_pr
        Pd_pr_Array(a) = Pd_pr
        Pt_pr_Array(a) = Pt_pr
        We_pr_Array(a) = We_pr
        Wa_pr_Array(a) = Wa_pr
        EFF_pr_Array(a) = ni_pr
        I_pr_Array(a) = I_pr
        V_pr_Array(a) = V

    End Sub

    'CALCOLO LA SOMMA DELLE AREE DEGLI UGELLI
    'OK
    Public Sub aggiorna_somma_d2_ugelli()

        somma_d2_ugelli = 0
        For index As Integer = 0 To (N_ugelli - 1)
            If ugelli(index).aperto Then somma_d2_ugelli = somma_d2_ugelli + Math.Pow(ugelli(index).diametro, 2)
        Next index

    End Sub

    'CALCOLO LA QMAX IN BASE ALL'ATTUALE CONFIGURAZIONE UGELLI
    Public Sub aggiorna_Qmax_da_configurazione_attuale_ugelli()

        Dim portata_max As UInteger = 0
        For indice = 0 To (N_ugelli - 1)
            If ugelli(indice).aperto Then portata_max = portata_max + ugelli(indice).Q_max
        Next

        Qmax = portata_max

    End Sub

    'aggiorno il numero che rappresenta la configurazione ugelli
    Public Sub aggiorna_num_conf_ugelli()

        Dim n_conf As UShort = 0

        For indice = 0 To (N_ugelli - 1)
            If ugelli(indice).aperto Then n_conf = n_conf + (2 ^ indice)
        Next

        Conf_Ug = n_conf

    End Sub

    'IMPOSTA LE VARIABILI DI CONTROLLO IN BASE ALLE IMPOSTAZIONI DELLA TAB "caratterizzazione sistema"
    'OK
    Public Sub imposta_parametri_controllo_da_testo()

        Tc_portata = CType(tb_Tc_portata.Text, Double)
        Kconfig_portata = CType(tb_Kconfig_portata.Text, Double)

        Tc_pressione = CType(tb_Tc_pressione.Text, Double)
        Kconfig_pressione = CType(tb_Kconfig_pressione.Text, Double)

    End Sub

    'AGGIORNA LA VISUALIZZAZIONE DEI PARAMETRI DI CONTROLLO

    Public Sub aggiorna_visualizzazione_parametri_controllo_attuali()

        lb_Tc_portata.Text = Tc_portata.ToString
        lb_Tc_pressione.Text = Tc_pressione.ToString
        lb_Kconfig_portata.Text = Kconfig_portata.ToString
        lb_Kconfig_pressione.Text = Kconfig_pressione.ToString
        lb_Kp_portata.Text = Kp_portata.ToString
        lb_Kp_pressione.Text = Kp_pressione.ToString
        lb_Kc_portata.Text = Kc_portata.ToString
        lb_Kc_pressione.Text = Kc_pressione.ToString

        lb_Tcampionamento.Text = Tcampionamento.ToString
        lb_Tp.Text = Tp.ToString
        lb_TETAp.Text = Tetap.ToString

    End Sub

    'DOPO AVER IMPOSTATO LA CONFIGURAZIONE UGELLI AGGIORNO I PARAMETRI DI CONTROLLO
    'QUESTI PARAMETRI SONO RICAVATI DA CARATTERIZZAZIONE SPERIMENTALE (vedi file "parametri_controllo.xls"
    'AD OGNI CAMBIO STRUTTURALE DELLA CAMERA (aggiunta ugelli ecc.) VA RIPETUTA LA CARATTERIZZAZIONE

    Public Sub aggiorna_parametri_controllo_da_configurazione_ugelli()

        Dim ascissa As Double = 1000 * somma_d2_ugelli

        Vaux_percent_max = Math.Round(1.3156 * Math.Pow(ascissa, 2) - 0.5912 * ascissa + 58.75, 1)
        If (Vaux_percent_max > 100) Then Vaux_percent_max = 100
        Kconfig_portata = Math.Round(-17.241 * Math.Pow(ascissa, 2) - 37.837 * ascissa + 1672.2)
        If (Kconfig_portata < 800) Then Kconfig_portata = 800

        Kconfig_pressione = Kconfig_pressione_01 'questo parametro è costante per tutte le configurazioni


    End Sub

    Private Sub aggiorna_parametri_controllo_effettivi()
        'in base alla configurazione del sistema stimo il suo guadagno
        Kp_portata = Kconfig_portata * somma_d2_ugelli
        'ricavo il guadagno del sistema in retroazione (Ti si pone uguale a Tp)
        Kc_portata = Tp / (Kp_portata * (Tetap + Tc_portata))

        'in base alla configurazione del sistema stimo il suo guadagno
        Kp_pressione = Kconfig_pressione * somma_d2_ugelli
        'ricavo il guadagno del sistema in retroazione (Ti si pone uguale a Tp)
        Kc_pressione = Tp / (Kp_pressione * (Tetap + Tc_pressione))
    End Sub

    'IMPOSTA LE VARIABILI DI CONTROLLO IN BASE ALLA CONFIGURAZIONE DEGLI UGELLI

    Private Sub aggiorna_parametri_da_configurazione_ugelli()

        If tutti_ugelli_chiusi() Then 'dovrebbe succedere solo durante il massima pressione
            Qmax = 1
            tb_Qmax_rif.Text = ""
            somma_d2_ugelli = 0

        Else             'ricalcolo tutti i parametri in base alla configurazione ugelli attuale

            aggiorna_somma_d2_ugelli()

            aggiorna_Qmax_da_configurazione_attuale_ugelli()

            aggiorna_num_conf_ugelli()

            aggiorna_parametri_controllo_da_configurazione_ugelli()

            aggiorna_parametri_controllo_effettivi()

            aggiorna_visualizzazione_parametri_controllo_attuali()

        End If

    End Sub

    'CONVERSIONE NUMERO CONFIGURAZIONE IN STRINGA BINARIA CORRISPONDENTE

    Public Function stringa_binaria_numconf(ByVal num_conf As UInt16) As String

        Dim stringa As String = Convert.ToString(num_conf, 2)

        If Len(stringa) < N_ugelli Then
            Dim caratteri_mancanti As Byte = N_ugelli - Len(stringa)
            For indice = 1 To caratteri_mancanti
                stringa = "0" & stringa
            Next
        End If

        Return stringa

    End Function


    'REGOLATORE PER L'INSEGUIMENTO DELLA PRESSIONE
    'OK
    Sub Insegui_Pressione(ByVal Read_press As Double, ByVal Target_press As Double, ByRef Press_Ok As Integer)

        Dim DVaux As Double
        Dim soglia As Double

        DVaux = 0

        'stimo l'errore attuale
        Err_act = Target_press - Read_press

        'stimo la variazione da applicare e imposto il ventilatore AUX
        DVaux = (Kc_pressione * (1 + Tcampionamento / Tp) * Err_act - Kc_pressione * Err_prev) 'DA VERIFICARE!!!!

        If IsNumeric(DVaux) Then Vaux_percent = Vaux_percent + DVaux
        Virgola(Vaux_percent, 2)

        If Vaux_percent > Vaux_percent_max Then Vaux_percent = Vaux_percent_max
        If Vaux_percent < Vaux_percent_min Then Vaux_percent = 0
        'i campi Vaux sull'interfaccia vengono aggiornati dopo l'aggiornamento dell'uscita aux_fan

        Err_prev = Err_act

        'se la portata è superiore ai 50Pa la soglia è 2Pa, altrimenti 1Pa
        If (Target_press > 50) Then
            soglia = 2
        Else
            soglia = 1
        End If

        If Abs(Err_act) < soglia Then
            Press_Ok = Press_Ok + 1
        Else
            Press_Ok = 0
        End If

    End Sub

    'REGOLATORE PER L'INSEGUIMENTO DELLA PORTATA
    'OK
    Sub Insegui_Portata(ByVal Read_q As Double, ByVal Target_q As Double, ByRef Q_Ok As Integer)

        Dim DVaux As Double
        Dim soglia As Double

        DVaux = 0

        'stimo l'errore attuale
        Err_act = Target_q - Read_q

        'stimo la variazione da applicare e imposto il ventilatore AUX
        DVaux = (Kc_portata * (1 + Tcampionamento / Tp) * Err_act - Kc_portata * Err_prev) 'DA VERIFICARE!!!!

        If IsNumeric(DVaux) Then Vaux_percent = Vaux_percent + DVaux
        Virgola(Vaux_percent, 2)

        If Vaux_percent > Vaux_percent_max Then Vaux_percent = Vaux_percent_max
        If Vaux_percent < Vaux_percent_min Then Vaux_percent = 0
        'i campi Vaux sull'interfaccia vengono aggiornati dopo l'aggiornamento dell'uscita aux_fan

        Err_prev = Err_act

        'se la portata è superiore ai 100m3/h la soglia è 3m3/h, altrimenti 1m3/h
        If (Target_q > 100) Then
            soglia = 3
        Else
            soglia = 1
        End If

        'se la portata è nei limiti incremento il contatore Q_Ok, che rappresenta da quanti cicli è ok la portata
        If Abs(Err_act) < soglia Then
            Q_Ok = Q_Ok + 1
        Else
            Q_Ok = 0
        End If

    End Sub

    'MEMORIZZAZIONE SU FOGLIO EXCEL DEI DATI DI INIZIO PROVA
    'OK
    Private Sub InizializzaReport(ByVal Template As String, ByVal FileDestinazione As String)

        Dim foglio_report, foglio_dati As String
        Dim xlFoglioDati As Microsoft.Office.Interop.Excel.Worksheet

        foglio_report = "report"
        foglio_dati = "@dati"

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.Visible = True
        xlBook = xlApp.Workbooks.Open(Template, , True)
        xlFoglioDati = xlBook.Worksheets(foglio_dati)
        xlSheet = xlBook.Worksheets(foglio_report)

        xlFoglioDati.Cells(6, 2).Value = tb_percorso_file_foto.Text 'percorso file foto

        xlSheet.Cells(12, 3).Value = Data  'data prova
        xlSheet.Cells(13, 3).Value = tb_esecutore.Text  'esecutore prova
        xlSheet.Cells(13, 6).Value = tb_rif_prova.Text  'riferimento numero prova

        xlSheet.Cells(16, 3).Value = tb_tipo_ventilatore.Text
        xlSheet.Cells(17, 3).Value = tb_produttore_ventilatore.Text
        xlSheet.Cells(18, 3).Value = tb_modello_ventilatore.Text
        xlSheet.Cells(19, 3).Value = tb_tensione_alimentazione.Text
        xlSheet.Cells(20, 3).Value = cb_frequenza_alimentazione.Text
        xlSheet.Cells(21, 3).Value = tb_diametro_eq_scarico.Text / 1000
        xlSheet.Cells(16, 6).Value = tb_produttore_motore.Text
        xlSheet.Cells(17, 6).Value = tb_codice_motore.Text
        xlSheet.Cells(18, 6).Value = tb_note_ventilatore.Text

        xlSheet.Cells(24, 3).Value = cb_tipo_prova.Text
        xlSheet.Cells(25, 3).Value = tb_note_test.Text
        xlSheet.Cells(25, 6).Value = cb_tipo_installazione.Text
        xlSheet.Cells(26, 6).Value = Pa / 100
        xlSheet.Cells(27, 6).Value = Hu
        xlSheet.Cells(28, 6).Value = Ta

        xlSheet.Cells(51, 17).Value = Roref


        'inserisco l'immagine
        xlApp.Run("inserisci_immagine")

        'salvo il file con il nome del report da creare
        xlBook.SaveAs(FileDestinazione, , , , , , , Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution)

        xlApp.WindowState = XlWindowState.xlMinimized

        xlFoglioDati = Nothing
        'xlApp.Quit()
        'xlSheet = Nothing
        'xlBook = Nothing
        'xlApp = Nothing

    End Sub

    'MEMORIZZAZIONE DATO SU FOGLIO EXCEL REPORT APERTO
    'usava gli array per popolare il report
    Sub MemorizzaDati_OLD(ByVal r As Integer)

        'dati grezzi
        xlSheet.Cells((r + 5), 9).Value = r + 1
        xlSheet.Cells((r + 5), 10).Value = Ro_pr_Array(r)
        xlSheet.Cells((r + 5), 11).Value = Qv_pr_Array(r)
        xlSheet.Cells((r + 5), 12).Value = Ps_pr_Array(r)
        xlSheet.Cells((r + 5), 13).Value = Pd_pr_Array(r)
        xlSheet.Cells((r + 5), 14).Value = Pt_pr_Array(r)
        xlSheet.Cells((r + 5), 15).Value = n_Array(r)
        xlSheet.Cells((r + 5), 16).Value = V_pr_Array(r)
        xlSheet.Cells((r + 5), 17).Value = I_pr_Array(r)
        xlSheet.Cells((r + 5), 18).Value = We_pr_Array(r)
        xlSheet.Cells((r + 5), 19).Value = Wa_pr_Array(r)
        xlSheet.Cells((r + 5), 20).Value = EFF_pr_Array(r)

        'dati corretti
        xlSheet.Cells((r + 30), 9).Value = r + 1
        xlSheet.Cells((r + 30), 10).Value = Ro_Array(r)
        xlSheet.Cells((r + 30), 11).Value = Qv_Array(r)
        xlSheet.Cells((r + 30), 12).Value = Ps_Array(r)
        xlSheet.Cells((r + 30), 13).Value = Pd_Array(r)
        xlSheet.Cells((r + 30), 14).Value = Pt_Array(r)
        xlSheet.Cells((r + 30), 15).Value = n_Array(r)
        xlSheet.Cells((r + 30), 16).Value = V_Array(r)
        xlSheet.Cells((r + 30), 17).Value = I_Array(r)
        xlSheet.Cells((r + 30), 18).Value = We_Array(r)
        xlSheet.Cells((r + 30), 19).Value = Wa_Array(r)
        xlSheet.Cells((r + 30), 20).Value = EFF_Array(r)

        xlBook.Save()

    End Sub


    'MEMORIZZAZIONE DATO SU FOGLIO EXCEL REPORT APERTO

    Sub MemorizzaDati(ByVal r As Integer)

        'dati grezzi
        xlSheet.Cells((r + 5), 9).Value = r + 1
        xlSheet.Cells((r + 5), 10).Value = Roa
        xlSheet.Cells((r + 5), 11).Value = qv_pr
        xlSheet.Cells((r + 5), 12).Value = Ps_pr
        xlSheet.Cells((r + 5), 13).Value = Pd_pr
        xlSheet.Cells((r + 5), 14).Value = Pt_pr
        xlSheet.Cells((r + 5), 15).Value = n
        xlSheet.Cells((r + 5), 16).Value = V_pr
        xlSheet.Cells((r + 5), 17).Value = I_pr
        xlSheet.Cells((r + 5), 18).Value = We_pr
        xlSheet.Cells((r + 5), 19).Value = Wa_pr
        xlSheet.Cells((r + 5), 20).Value = ni_pr

        'dati corretti
        xlSheet.Cells((r + 30), 9).Value = r + 1
        xlSheet.Cells((r + 30), 10).Value = Roref
        xlSheet.Cells((r + 30), 11).Value = qv
        xlSheet.Cells((r + 30), 12).Value = Ps
        xlSheet.Cells((r + 30), 13).Value = Pd
        xlSheet.Cells((r + 30), 14).Value = Pt
        xlSheet.Cells((r + 30), 15).Value = n
        xlSheet.Cells((r + 30), 16).Value = V
        xlSheet.Cells((r + 30), 17).Value = I
        xlSheet.Cells((r + 30), 18).Value = We
        xlSheet.Cells((r + 30), 19).Value = Wa
        xlSheet.Cells((r + 30), 20).Value = ni


        xlBook.Save()

    End Sub

    Public Sub FinalizzaReport()

        xlBook.Save()
        'xlApp.Quit()

        xlSheet = Nothing
        xlBook = Nothing
        xlApp = Nothing

    End Sub

    'MEMORIZZAZIONE ERRORI
    'OK
    Public Sub PopolaDB_error(ByVal Device_Error As String)

        'MessageBox.Show(Device_Error)

        lb_errore.Text = Device_Error
        lb_errore.BackColor = Color.Red

        'tabella_errori.Rows.Add()
        'tabella_errori.BeginEdit(True)
        'tabella_errori.Rows.GetLastRow(DataGridViewElementStates.Displayed)
        ' tabella_errori.CurrentCell.Value = Device_Error
        ' tabella_errori.EndEdit()

    End Sub

    'INIZIALIZZAZIONE VARIABILI E INIBIZIONE PULSANTI E SELEZIONI ALLO START DELLA PROVA AUTOMATICA
    'VERIFICARE REPORT
    Sub Start_Prova_Auto()

        '  Dim stringr As String
        lb_messaggio.Text = "INIZIO TEST AUTOMATICO"
        Timer_idle.Stop()

        'INIZIALIZZAZIONE VARIABILI
        StandBy = 0
        Limit = 0
        Press_Target_Ok = 0
        Q_Target_Ok = 0

        Buffer_Dp = array_3_vuoto
        Buffer_I_p = array_3_vuoto
        Buffer_P_cam1 = array_3_vuoto
        Buffer_P_cam2 = array_3_vuoto
        Buffer_TACHO = array_3_vuoto
        Buffer_V_p = array_3_vuoto
        Buffer_We_p = array_3_vuoto

        Fase = 0
        DQ = 0
        'Ugelli_prev = Ugello
        'Ugelli_act = Ugello

        'inizialmente imposto 10 punti (9 portata + max press)
        'in seguito alla misura della portata massima, aggiorno il numero dei punti in base alle impostazioni dell'utente
        num_totale_punti = 10

        nq = 0 'punto iniziale a massima portata

        Ps_Array = array_20_vuoto
        Qv_Array = array_20_vuoto
        Ro_Array = array_20_vuoto
        Pd_Array = array_20_vuoto
        Pt_Array = array_20_vuoto
        n_Array = array_20_vuoto
        I_Array = array_20_vuoto
        We_Array = array_20_vuoto
        Wa_Array = array_20_vuoto
        EFF_Array = array_20_vuoto
        V_Array = array_20_vuoto

        Ps_pr_Array = array_20_vuoto
        Qv_pr_Array = array_20_vuoto
        Ro_pr_Array = array_20_vuoto
        Pd_pr_Array = array_20_vuoto
        Pt_pr_Array = array_20_vuoto
        We_pr_Array = array_20_vuoto
        Wa_pr_Array = array_20_vuoto
        EFF_pr_Array = array_20_vuoto
        I_pr_Array = array_20_vuoto
        V_pr_Array = array_20_vuoto

        Err_prev = 0
        Err_act = 0

        'leggo la stazione barometrica per avere i dati da mettere nel file
        Ok_Barometro = stazione_barometrica.LeggiDati()
        Ta = stazione_barometrica.temperatura
        Hu = stazione_barometrica.umidità
        Pa = stazione_barometrica.pressione
        lb_Ta.Text = Ta.ToString
        lb_Hu.Text = Hu.ToString
        lb_Pa.Text = Pa.ToString

        aggiorna_parametri_da_configurazione_ugelli()

        'accendo il ventilatore ausiliario
        Vaux_percent = Vaux_percent_min
        aggiorna_uscita_Vaux()
        serranda_percent = serranda_min


        'DA VERIFICARE!!!!!!
        InizializzaReport(tb_template_report.Text, tb_percorso_file_report.Text)

        Timer_prova_automatica.Interval = Tcampionamento * 1000
        Timer_prova_automatica.Enabled = 1
        Timer_prova_automatica.Start()

    End Sub

    'RIPRISTINO INTERFACCIA (ENTRAMBI I CASI)
    'OK
    Private Sub ripristina_interfaccia_fine_prova()

        'riabilito i controlli non utilizzabili durante la prova

        'scheda test
        bt_azzera_pressioni.Enabled = True
        bt_dati_aria_default.Enabled = True
        bt_calcola_rho.Enabled = True

        bt_start_test_auto.Enabled = True
        bt_start_test_man.Enabled = True

        tb_rho_rif.Enabled = True
        tb_Ta_aria_rif.Enabled = True
        tb_Hu_aria_rif.Enabled = True
        tb_Pb_aria_rif.Enabled = True

        tb_percorso_file_foto.Enabled = True
        tb_percorso_file_report.Enabled = True
        pb_percorso_file_foto.Enabled = True
        pb_percorso_file_report.Enabled = True

        lb_raggiungimento_punto.Visible = False
        pb_raggiungimento_punto.Visible = False
        bt_fine_test.Enabled = False
        bt_max_press_termina_test.Enabled = False

        'scheda impostazioni test
        bt_genera_configurazione_ugelli.Enabled = True
        cb_tipo_installazione.Enabled = True
        cb_tipo_prova.Enabled = True

    End Sub

    'RIPRISTINO INTERFACCIA PROVA AUTOMATICA
    'OK
    Private Sub ripristina_interfaccia_prova_automatica()

        ch_vaux.Enabled = True
        ch_portata.Enabled = True
        ch_pressione.Enabled = True


    End Sub

    'RIPRISTINO INTERFACCIA PROVA MANUALE
    'OK
    Private Sub ripristina_interfaccia_prova_manuale()

        bt_salva_punto.Enabled = False

    End Sub

    'ARRESTO PROVA, STOP VENTILATORE AUX E RIPRISTINO PULSANTI E SELEZIONI DI CONFIGURAZIONE
    'OK
    Sub Stop_prova()

        Vaux_percent = 0
        serranda_percent = 0
        Ok_Analog = aux_fan.azzera_uscita()
        Ok_Analog = aux_fan.imposta_percentuale_serranda(0)
        lb_Vaux_mandata.Text = "0"
        lb_Vaux_aspirazione.Text = "0"
        lb_percent_serranda.Text = "0"

        lb_messaggio.Text = "FINALIZZAZIONE PROVA..."
        pb_avanzamento_test.Value = 100

        ripristina_interfaccia_fine_prova()

        If Test_Manuale Then
            Test_Manuale = False
            Timer_prova_manuale.Stop()
            ripristina_interfaccia_prova_manuale()
        Else
            Timer_prova_automatica.Stop()
            ripristina_interfaccia_prova_automatica()
        End If

        FinalizzaReport()

        Fase = 0
        nq = 0

        Timer_idle.Start()
        lb_messaggio.Text = "STANDBY"
        pb_avanzamento_test.Value = 0

        Beep()
        Beep()

    End Sub

    'CALCOLO DEL PUNTO A PRESSIONE MAX E A SEGUIRE INTERRUZIONE PROVA
    'OK
    Private Sub bt_max_press_termina_test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_max_press_termina_test.Click
        'in questo caso conferma l'acquisizione della pressione massima
        Stop_Max_Press = True

    End Sub

    'DECISIONE PUNTI DA TESTARE A SECONDA DELLA SELEZIONE DELL'UTENTE
    'OK - DA ESPANDERE
    Public Sub PopolaArrayPunti()

        array_punti_portata = array_20_vuoto

        'N punti equidistanti (default)
        num_totale_punti = CType(tb_num_punti_portata.Text, Integer)
        DQ = qv_max / num_totale_punti
        Virgola(DQ, 1)
        For indice As Integer = 0 To num_totale_punti - 1
            array_punti_portata(indice) = DQ * (num_totale_punti - indice)
        Next indice

        'eventualmente includo il punto a massima pressione
        If ch_misura_portata_massima.Checked Then
            num_totale_punti = num_totale_punti + 1
            array_punti_portata(num_totale_punti - 1) = 0
        End If


    End Sub

    Private Sub passa_alla_fase_2()
        'passo alla fase 2
        Fase = 2
        StandBy = 0
        TestMaxPress = 0
        Stop_Max_Press = False

        'SPENGO IL VENTILATORE AUX e APRO LA SERRANDA
        Ok_Analog = aux_fan.azzera_uscita()
        lb_Vaux_mandata.Text = "0"
        lb_Vaux_aspirazione.Text = "0"

        Ok_Analog = aux_fan.imposta_percentuale_serranda(0)
        lb_percent_serranda.Text = "0"

    End Sub

    Private Sub acquisisci_max_press()

        'RICERCA PUNTO A MASSIMA PRESSIONE

        pb_avanzamento_test.Value = 95
        If (lb_messaggio.Text <> "ATTENDERE PREGO.") Then
            lb_messaggio.Text = "RICERCA DI MASSIMA PRESSIONE, CHIUDERE TUTTI GLI UGELLI E PREMERE OK."
        End If

        StandBy = 1
        bt_status.Visible = True

        riempi_buffer_circolari(TestMaxPress)

        If (TestMaxPress > 0) Then 'dopo che l'utente ha premuto OK inizio a contare i cicli di stabilità

            If (TestMaxPress > num_cicli_lettura_stabile) Then

                StandBy = 0
                bt_status.Visible = False
                bt_status.Enabled = True

                PopolaArray(nq)
                media_buffer_circolari() 'calcola i valori medi acquisiti
                esegui_calcoli()
                MemorizzaDati(nq)
                Beep()

                Fase = 3
            Else
                'aggiorno la progress bar
                ratio_raggiungimento_punto = TestMaxPress / num_cicli_lettura_stabile
                If ratio_raggiungimento_punto > 1 Then ratio_raggiungimento_punto = 1
                pb_raggiungimento_punto.Value = CType(ratio_raggiungimento_punto * 100, Integer)

            End If

            TestMaxPress = TestMaxPress + 1

        End If

    End Sub

    'CLOCK DI LETTURA DATI E GESTIONE INSEGUIMENTI
    'OK
    Private Sub Timer_prova_automatica_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_prova_automatica.Tick

        esegui_ciclo_di_acquisizione()

        tb_fase.Text = Fase.ToString

        Select Case Fase

            Case 0 'PRIMO PUNTO A PORTATA MASSIMA

                pb_avanzamento_test.Value = 5
                lb_messaggio.Text = "RICERCA PORTATA MASSIMA"
                Insegui_Pressione(Ps, 0, Press_Target_Ok)

                'aggiorno la progress bar
                ratio_raggiungimento_punto = Press_Target_Ok / num_cicli_lettura_stabile
                If ratio_raggiungimento_punto > 1 Then ratio_raggiungimento_punto = 1
                pb_raggiungimento_punto.Value = CType(ratio_raggiungimento_punto * 100, Integer)

                aggiorna_uscita_Vaux()

                If ((Vaux_percent = Vaux_percent_max) And (Ps > 0.9)) Then Limit = Limit + 1

                'SE HO RAGGIUNTO IL PUNTO ACQUISISCO IL DATO E PASSO ALLA FASE 1
                If (Press_Target_Ok > num_cicli_lettura_stabile) Then

                    Virgola(qv, 1)
                    qv_max = qv
                    PopolaArrayPunti()

                    'inserisco il punto corrente nell'array punti
                    PopolaArray(nq)
                    Ps_Array(nq) = 0
                    Ps_pr_Array(nq) = 0
                    Pt_Array(nq) = Pd_Array(nq)
                    Pt_pr_Array(nq) = Pd_pr_Array(nq)

                    media_buffer_circolari() 'calcola i valori medi acquisiti
                    esegui_calcoli()
                    MemorizzaDati(nq)

                    Limit = 0
                    Fase = 1

                    nq = nq + 1
                    Press_Target_Ok = 0
                    Beep()

                Else
                    If Stop_Max_Press Then 'se ho forzato l'acquisizione del punto di massima pressione passo subito alla fase 2
                        passa_alla_fase_2()
                    Else 'altrimenti leggo normalmente i dati
                        riempi_buffer_circolari(Press_Target_Ok)
                    End If

                End If

                'SE IL PUNTO A MASSIMA PORTATA NON E' RAGGIUNGIBILE ANNULLO LA PROVA
                If (Limit > num_cicli_lettura_stabile) Then
                    lb_messaggio.Text = "CONFIGURAZIONE TROPPO CHIUSA. CAMBIARE CONFIGURAZIONE E RIPETERE IL TEST"
                    Timer_prova_automatica.Stop()
                    'Timer_prova_manuale.Stop()
                    Timer_annullamento_prova.Interval = 3000
                    Timer_annullamento_prova.Enabled = 1
                    Timer_annullamento_prova.Start()
                    Beep()
                    Beep()
                End If

            Case 1 'RICERCA PUNTI ALLE PORTATE SCAGLIONATE

                pb_avanzamento_test.Value = 100 * nq / num_totale_punti

                lb_messaggio.Text = "RICERCA PUNTO A PORTATA " & array_punti_portata(nq) & " M3/H"
                Insegui_Portata(qv, array_punti_portata(nq), Q_Target_Ok)

                'aggiorno la progress bar
                ratio_raggiungimento_punto = Q_Target_Ok / num_cicli_lettura_stabile
                If ratio_raggiungimento_punto > 1 Then ratio_raggiungimento_punto = 1
                pb_raggiungimento_punto.Value = CType(ratio_raggiungimento_punto * 100, Integer)

                aggiorna_uscita_Vaux()

                If ((Vaux_percent = 0) And (nq < (num_totale_punti - 1))) Then Limit = Limit + 1

                'SE HO RAGGIUNTO IL PUNTO ACQUISISCO IL DATO
                If (Q_Target_Ok > num_cicli_lettura_stabile) Then

                    PopolaArray(nq)
                    media_buffer_circolari() 'calcola i valori medi acquisiti
                    esegui_calcoli()
                    MemorizzaDati(nq)
                    nq = nq + 1
                    Q_Target_Ok = 0
                    Beep()

                Else
                    riempi_buffer_circolari(Q_Target_Ok)
                End If

                'CONDIZIONI DI PASSAGGIO AL PUNTO A MAX PRESS
                If ((array_punti_portata(nq) = 0) Or (Stop_Max_Press)) Then
                    passa_alla_fase_2()
                End If

                ''SE IL DATO NON E' RAGGIUNGIBILE CONSIGLIO IL CAMBIO CONFIGURAZIONE UGELLI
                If (Limit > num_cicli_lettura_stabile) Then

                    If (Conf_Ug = 1) Then
                        'se sono nella configurazione minima non posso scendere di portata, per cui forzo la misura della pressione massima

                        Stop_Max_Press = True

                    Else    'se sono in un altra configurazione 

                        lb_messaggio.Text = "ADEGUARE CONFIGURAZIONE UGELLI PER " & array_punti_portata(nq) & " M3/H " & "E POI PREMERE OK"

                        'SUGGERISCO LA CONFIGURAZIONE UGELLI ADEGUATA
                        tb_Qmax_rif.Text = array_punti_portata(nq).ToString
                        stima_configurazione_ugelli()

                        'SPENGO IL VENTILATORE AUX
                        Ok_Analog = aux_fan.azzera_uscita
                        lb_Vaux_mandata.Text = Vaux_percent
                        lb_Vaux_aspirazione.Text = Vaux_percent

                        Timer_prova_automatica.Stop()
                        bt_status.Visible = True
                        bt_status.Enabled = True
                        Beep()
                        Beep()
                    End If

                End If

                'SE LA PRESSIONE VA FUORI SCALA ANNULLO LA PROVA
                If (Vaux_percent > Vaux_percent_max) Then StandBy = 0

                If (((cb_tipo_prova.Text = "Mandata") And (Abs(P_cam1) >= Pmax_assoluta)) Or _
                     ((cb_tipo_prova.Text = "Aspirazione") And (Abs(P_cam2) >= Pmax_assoluta))) Then
                    lb_messaggio.Text = "VENTILATORE CON PRESSIONE TROPPO ELEVATA. PROVA TERMINATA."

                    Timer_prova_automatica.Stop()
                    Fase = 0
                    nq = 0

                    Timer_annullamento_prova.Interval = 3000
                    Timer_annullamento_prova.Enabled = 1
                    Timer_annullamento_prova.Start()

                    Beep()
                    Beep()

                End If

            Case 2

                acquisisci_max_press()

            Case 3

                'FINE PROVA
                Stop_prova()

            Case 4

            Case Else

        End Select

    End Sub

    'INPUT DA UTENTE PER PROSEGUIRE CON LA PROVA NEI CASI DI CONFIGURAZIONE UGELLI AGGIORNATA E TUTTO CHIUSO
    'OK
    Private Sub bt_status_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_status.Click

        Select Case Fase
            Case 1 'FASE 1: PUNTI IN PORTATA - se ho premuto significa che ho cambiato configurazione
                Limit = 0
                bt_status.Visible = False

                Vaux_percent = Vaux_percent_min
                aux_fan.imposta_percentuale(Vaux_percent)
                lb_Vaux_mandata.Text = Vaux_percent
                lb_Vaux_aspirazione.Text = Vaux_percent

                Err_prev = 0
                Err_act = 0
                Q_Target_Ok = 0
                Timer_prova_automatica.Start()

            Case 2 'FASE 2: MISURA MAX PRESS - se ho premuto significa che ho chiuso tutti gli ugelli
                StandBy = 0
                TestMaxPress = 1
                lb_messaggio.Text = "ATTENDERE PREGO."
                bt_status.Enabled = False
            Case Else
        End Select

    End Sub


    'RITARDO STOP PROVA
    'OK
    Private Sub Timer_annullamento_prova_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_annullamento_prova.Tick

        Timer_annullamento_prova.Stop()
        Stop_prova()

    End Sub

    ' START DELLA PROVA MANUALE
    'DA VERIFICARE REPORT
    Private Sub Start_Prova_Man()

        'Dim stringr As String
        ' Dim j As Integer

        lb_messaggio.Text = "TEST MANUALE IN CORSO"
        Timer_idle.Stop()

        'INIZIALIZZAZIONE VARIABILI
        num_totale_punti = num_max_punti
        nq = 0
        'j = 0
        Fase = 0

        Buffer_Dp = array_3_vuoto
        Buffer_I_p = array_3_vuoto
        Buffer_P_cam1 = array_3_vuoto
        Buffer_P_cam2 = array_3_vuoto
        Buffer_TACHO = array_3_vuoto
        Buffer_V_p = array_3_vuoto
        Buffer_We_p = array_3_vuoto

        Ps_Array = array_20_vuoto
        Qv_Array = array_20_vuoto
        Ro_Array = array_20_vuoto
        Pd_Array = array_20_vuoto
        Pt_Array = array_20_vuoto
        n_Array = array_20_vuoto
        I_Array = array_20_vuoto
        We_Array = array_20_vuoto
        Wa_Array = array_20_vuoto
        EFF_Array = array_20_vuoto
        V_Array = array_20_vuoto

        Ps_pr_Array = array_20_vuoto
        Qv_pr_Array = array_20_vuoto
        Ro_pr_Array = array_20_vuoto
        Pd_pr_Array = array_20_vuoto
        Pt_pr_Array = array_20_vuoto
        We_pr_Array = array_20_vuoto
        Wa_pr_Array = array_20_vuoto
        EFF_pr_Array = array_20_vuoto
        I_pr_Array = array_20_vuoto
        V_pr_Array = array_20_vuoto


        'leggo la stazione barometrica per avere i dati da mettere nel file
        Ok_Barometro = stazione_barometrica.LeggiDati()
        Ta = stazione_barometrica.temperatura
        Hu = stazione_barometrica.umidità
        Pa = stazione_barometrica.pressione
        lb_Ta.Text = Ta.ToString
        lb_Hu.Text = Hu.ToString
        lb_Pa.Text = Pa.ToString

        'valori iniziale di portata e pressione per non far partire a canna il ventilatore aux
        Qman = 1
        Pman = 500

        aggiorna_parametri_da_configurazione_ugelli()
        'accendo il ventilatore ausiliario
        aux_fan.abilita()
        'accendo la serranda
        aux_fan.abilita_uscita_serranda()

        'DA VERIFICARE!!!!!!
        InizializzaReport(tb_template_report.Text, tb_percorso_file_report.Text)

        Timer_prova_manuale.Interval = Tcampionamento * 1000
        Timer_prova_manuale.Enabled = 1
        Timer_prova_manuale.Start()

    End Sub

    'LETTURA DI TUTTI GLI STRUMENTI DURANTE LA PROVA
    'OK
    Private Sub leggi_misurazioni()

        'divido le letture del wattmetro perché se lo si interroga più di una volta al secondo si offende

        Select Case ciclo_lettura

            Case 1
                resetta_label_errori()

                'nei ciclo 1 leggo stazione barometrica (400ms)

                Ok_Barometro = stazione_barometrica.LeggiDati()
                Ta = stazione_barometrica.temperatura
                Hu = stazione_barometrica.umidità
                Pa = stazione_barometrica.pressione
                T_ing_ug = Ta
                T_cam2 = Ta

                'e la tensione (100ms)

                If rb_infratek.Checked Then 'leggo dal wattmetro selezionato
                    V_p = wattmetro.tensione_V(Ok_Wattmetro)
                Else
                    V_p = esam.tensione(Ok_Wattmetro)
                End If


                If Ok_Wattmetro Then
                    V_p_prev = V_p
                Else
                    V_p = V_p_prev
                End If

            Case 2
                'nel ciclo 2 la corrente (100ms)
                Select Case cb_frequenza_alimentazione.Text
                    Case "DC"
                        I_p = multimetro.valore_letto(3)

                    Case Else

                        If rb_infratek.Checked Then 'leggo dal wattmetro selezionato
                            I_p = wattmetro.corrente_A(Ok_Wattmetro)
                        Else
                            I_p = esam.corrente(Ok_Wattmetro)
                        End If

                        If Ok_Wattmetro Then
                            I_p_prev = I_p
                        Else
                            I_p = I_p_prev
                        End If

                End Select


            Case 3
                'nel ciclo 3 la potenza (100ms)
                Select Case cb_frequenza_alimentazione.Text
                    Case "DC"
                        We_p = V * I_p
                    Case Else

                        If rb_infratek.Checked Then 'leggo dal wattmetro selezionato
                            We_p = wattmetro.potenza_W(Ok_Wattmetro)
                        Else
                            We_p = esam.potenza(Ok_Wattmetro)
                        End If

                        If Ok_Wattmetro Then
                            We_p_prev = We_p
                        Else
                            We_p = We_p_prev
                        End If
                End Select

                ciclo_lettura = 0

            Case Else
                ciclo_lettura = 0

        End Select

        ciclo_lettura = ciclo_lettura + 1

        'tachimetro (tempo trascurabile)
        n = tachimetro.rpm

        'pressioni (100ms cadauna)
        If Fase = 2 Then 'se sto misurando la pressione max
            Dp = 0
        Else
            Dp = trasm_DP.pressione_Pa(Ok_DP)
            If Ok_DP Then
                Virgola(Dp, 1)
                dp_prev = Dp
            Else
                Dp = dp_prev
            End If
        End If

        Ptemp = trasm_P1.pressione_Pa(Ok_P1)
        If Ok_P1 Then
            Virgola(Ptemp, 1)
            Ptemp_prev = Ptemp
        Else
            Ptemp = Ptemp_prev
        End If

        Select Case cb_tipo_prova.Text
            Case "Mandata"
                P_cam1 = Ptemp
                P_ing_ug = P_cam1
            Case "Aspirazione"
                P_cam2 = Ptemp
                P_ing_ug = P_cam2 + Dp
            Case Else
        End Select


    End Sub

    'ESECUZIONE MEDIA SU 3 VALORI DELLE GRANDEZZE MAGGIORMENTE INSTABILI
    'OK
    Private Sub media_misurazioni()

        'pressioni e numero giri
        n = Filtro_Dati(n, Buffer_TACHO, step_buf_TACHO)
        Virgola(n, 0)
        Dp = Filtro_Dati(Dp, Buffer_Dp, step_buf_Dp)
        Virgola(Dp, 1)

        Select Case cb_tipo_prova.Text
            Case "Mandata"
                P_cam1 = Filtro_Dati(P_cam1, Buffer_P_cam1, step_buf_P_cam1)
                Virgola(P_cam1, 1)
                P_ing_ug = P_cam1

            Case "Aspirazione"
                P_cam2 = Filtro_Dati(P_cam2, Buffer_P_cam2, step_buf_P_cam2)
                Virgola(P_cam2, 1)
                P_ing_ug = P_cam2 + Dp

            Case Else
        End Select

        'grandezze elettriche
        'V_p = Filtro_Dati(V_p, Buffer_V_p, step_buf_V_p)
        Virgola(V_p, 1)
        ' We_p = Filtro_Dati(We_p, Buffer_We_p, step_buf_We_p)
        Virgola(We_p, 1)
        ' I_p = Filtro_Dati(I_p, Buffer_I_p, step_buf_I_p)
        Virgola(I_p, 3)

    End Sub

    'IMPOSTA L'USCITA ANALOGICA E I CONTROLLI DEL VENTILATORE AUSILIARIO
    'OK
    Private Sub aggiorna_uscita_Vaux()

        aux_fan.imposta_percentuale(Vaux_percent)
        lb_Vaux_mandata.Text = Vaux_percent.ToString
        lb_Vaux_aspirazione.Text = Vaux_percent.ToString

    End Sub

    'IMPOSTA L'USCITA ANALOGICA E I CONTROLLI DELLA SERRANDA
    'OK
    Private Sub aggiorna_uscita_serranda()

        aux_fan.imposta_percentuale_serranda(serranda_percent)
        lb_percent_serranda.Text = serranda_percent.ToString

    End Sub


    '**************************************************************************************
    '                          PROCEDURE DI CALCOLO DOPO LE MISURE
    '**************************************************************************************

    'CALCOLO SOMMATORIA DI Cj * dj^2 /4
    'OK
    Private Sub calcola_sum_cj_d2()

        Dim index, iter As Integer
        Dim Cj_prev(N_ugelli - 1), Cj_act(N_ugelli - 1), Redj(N_ugelli - 1), Ro As Double

        Select Case cb_tipo_prova.Text
            Case "Mandata"
                Ro = Ro6
            Case "Aspirazione"
                Ro = Ro7
            Case Else
        End Select

        index = 0
        Sum_Cj_d2 = 0
        For indice = 0 To (N_ugelli - 1)
            Cj_prev(indice) = 0.95
            Cj_act(indice) = 0
            Redj(indice) = 0
        Next

        For index = 0 To (N_ugelli - 1)
            iter = 0
            If ugelli(index).aperto Then
                Do Until ((Math.Abs(Cj_prev(index) - Cj_act(index))) < 0.001) Or (iter > 2000)
                    If iter <> 0 Then
                        Cj_prev(index) = Cj_act(index)
                    End If
                    Redj(index) = 1000000 * epsilon * Cj_prev(index) * ugelli(index).diametro * (Math.Sqrt(2 * Ro * Dp)) / (17.1 + 0.048 * T_ing_ug)
                    Cj_act(index) = 0.9986 + ugelli(index).k1_alfa / (Math.Sqrt(Redj(index))) + ugelli(index).k1_alfa / Redj(index)
                    iter = iter + 1
                Loop
                Sum_Cj_d2 = Sum_Cj_d2 + (Cj_act(index) * Math.Pow(ugelli(index).diametro, 2) / 4)

            End If
        Next


    End Sub

    'CALCOLI INTERMEDI DA LETTURE DATI (VALIDI PER TUTTI I CASI)
    'OK
    Sub Calcoli_intermedi_generici()

        Psat_Ta = 610.8 + 44.442 * Ta + 1.4133 * Math.Pow(Ta, 2) + 0.02768 * Math.Pow(Ta, 3) + 0.000255667 * Math.Pow(Ta, 4) + _
            0.00000289166 * Math.Pow(Ta, 5)
        Pv = Hu * Psat_Ta / 100
        Teta_a = Ta + 273.15
        Roa = (Pa - 0.378 * Pv) / (287 * Teta_a)
        Virgola(Roa, 3)
        Rw = Pa / (Teta_a * Roa)
        Ro1 = Roa

    End Sub

    'CALCOLI INTERMEDI DA LETTURE DATI (VALIDI PER PROVA IN MANDATA)
    'OK
    Sub Calcoli_intermedi_mandata()

        P6 = Pa + P_ing_ug
        P4 = Pa + P_cam1
        Teta6 = T_ing_ug + 273.15
        Ro6 = P6 / (Teta6 * Rw)
        calcola_sum_cj_d2()
        qm = coeff_TUV_portata * epsilon * 3.14159 * (Math.Sqrt(2 * Ro6 * Dp)) * Sum_Cj_d2
        If ((Not IsNumeric(qm)) Or (Fase = 2)) Then qm = 0


    End Sub

    'CALCOLI INTERMEDI DA LETTURE DATI (VALIDI PER PROVA IN ASPIRAZIONE)
    'OK
    Sub Calcoli_intermedi_aspirazione()

        P7 = Pa + P_ing_ug
        Psg3 = Pa + P_cam2
        Teta7 = T_ing_ug + 273.15
        Ro7 = P7 / (Teta7 * Rw)
        Teta3 = T_cam2 + 273.15
        Ro3 = Psg3 / (Teta3 * Rw)
        calcola_sum_cj_d2()
        qm = coeff_TUV_portata * epsilon * 3.14159 * (Math.Sqrt(2 * Ro7 * Dp)) * Sum_Cj_d2
        If ((Not IsNumeric(qm)) Or (Fase = 2)) Then qm = 0

    End Sub

    'CALCOLO RISULTATI TEST ALLE CONDIZIONI DI PROVA (VALIDI PER PROVA IN MANDATA)
    'OK
    Sub Risultati_Mandata_prova()

        'portata, pressione statica e dinamica
        'la correzione TUV è già stata applicata alla portata massica
        qv_pr = 3600 * qm / Ro6
        Ps_pr = coeff_TUV_pressione * P_cam1
        Pd_pr = Math.Pow((qm / A2), 2) / (2 * Ro6)

    End Sub

    'CALCOLO RISULTATI TEST ALLE CONDIZIONI DI PROVA (VALIDI PER PROVA IN ASPIRAZIONE)
    'OK
    Sub Risultati_Aspirazione_prova()

        'portata, pressione statica e dinamica
        qv_pr = 3600 * qm / Ro7
        Ps_pr = -coeff_TUV_pressione * P_cam2
        Pd_pr = Math.Pow((qm / A2), 2) / (2 * Ro7)

    End Sub

    'CALCOLO RISULTATI TEST ALLE CONDIZIONI DI PROVA (VALIDI PER ENTRAMBE LE CONFIGURAZIONI)
    'OK
    Sub Risultati_Comuni_prova()

        'approssimo le grandezze già calcolate
        Virgola(qv_pr, 1)
        Virgola(Ps_pr, 1)
        Virgola(Pd_pr, 1)

        'pressione totale
        Pt_pr = Ps_pr + Pd_pr

        'grandezze elettriche
        We_pr = We_p
        I_pr = I_p
        V_pr = V_p

        'potenza aeraulica
        If ((Install = "A") Or (Install = "C")) Then
            Wa_pr = Ps_pr * qv_pr / 3600
        End If
        If ((Install = "B") Or (Install = "D")) Then
            Wa_pr = Pt_pr * qv_pr / 3600
        End If
        'efficienza
        ni_pr = Wa_pr / We_pr * 100
        Virgola(Wa_pr, 1)
        Virgola(ni_pr, 1)

    End Sub

    'CALCOLO RISULTATI TEST ALLE CONDIZIONI DI RIFERIMENTO (VALIDI PER PROVA IN MANDATA)
    'OK
    Sub Risultati_Mandata_corretti()

        'tensione di alimentazione
        If ch_correzione_tensione_alimentazione.Checked Then
            V = CType(tb_tensione_alimentazione.Text, Double)
        Else
            V = V_pr
        End If

        'pressione statica
        Ps = Ps_pr * Roref / Ro6
        Virgola(Ps, 1)

        'corrente
        'I = I_p * Roref * V * V / (Ro6 * V_pr * V_pr)
        I = I_p * V / V_pr
        Virgola(I, 3)
        'potenza elettrica
        'We = We_p * Roref * V * V / (Ro6 * V_pr * V_pr)
        We = We_p * V * V / (V_pr * V_pr)
        'dopo la prova di confronto con TUV, i parametri elettrici non vengono più corretti con la densità


    End Sub

    'CALCOLO RISULTATI TEST ALLE CONDIZIONI DI RIFERIMENTO (VALIDI PER PROVA IN ASPIRAZIONE)
    'OK
    Sub Risultati_Aspirazione_corretti()

        'tensione di alimentazione
        If ch_correzione_tensione_alimentazione.Checked Then
            V = CType(tb_tensione_alimentazione.Text, Double)
        Else
            V = V_pr
        End If

        'pressione statica
        Ps = Ps_pr * Roref / Ro3
        Virgola(Ps, 1)

        'corrente
        'I = I_p * Roref * V * V / (Ro3 * V_pr * V_pr)
        I = I_p * V / V_pr
        Virgola(I, 3)
        'potenza elettrica
        'We = We_p * Roref * V * V / (Ro3 * V_pr * V_pr)
        We = We_p * V * V / (V_pr * V_pr)
        'dopo la prova di confronto con TUV, i parametri elettrici non vengono più corretti con la densità


    End Sub

    'CALCOLO RISULTATI TEST ALLE CONDIZIONI DI RIFERIMENTO (VALIDI PER TUTTI I CASI)
    'OK
    Sub Risultati_Comuni_corretti()

        'portata
        qv = 3600 * qm / Roref
        Virgola(qv, 1)
        'pressione dinamica
        Pd = Math.Pow((qm / A2), 2) / (2 * Roref)
        Virgola(Pd, 1)
        'pressione totale
        Pt = Ps + Pd

        'potenza aeraulica
        If ((Install = "A") Or (Install = "C")) Then
            Wa = Ps * qv / 3600
        End If
        If ((Install = "B") Or (Install = "D")) Then
            Wa = Pt * qv / 3600
        End If

        'rendimento
        ni = Wa / We * 100
        Virgola(ni, 1)

        'dopo i calcoli approssimo le potenze al decimo di W
        Virgola(We, 1)
        Virgola(Wa, 1)

    End Sub

    'ESECUZIONE TUTTI I CALCOLI DOPO LE MISURE
    'OK
    Private Sub esegui_calcoli()

        Calcoli_intermedi_generici()

        Select Case cb_tipo_prova.Text
            Case "Mandata"
                Calcoli_intermedi_mandata()

                Risultati_Mandata_prova()
                Risultati_Comuni_prova()

                Risultati_Mandata_corretti()
                Risultati_Comuni_corretti()

            Case "Aspirazione"
                Calcoli_intermedi_aspirazione()

                Risultati_Aspirazione_prova()
                Risultati_Comuni_prova()

                Risultati_Aspirazione_corretti()
                Risultati_Comuni_corretti()

            Case Else
        End Select

    End Sub


    '*********************************************************************************
    ' TERMINE PROCEDURE CALCOLO DOPO MISURE
    '***********************************************************************************

    'AGGIORNAMENTO CAMPI DELL'INTERFACCIA CON I VALORI LETTI E SALVATI NELLE VARIABILI
    Private Sub aggiorna_interfaccia_con_letture()

        tb_pressione_statica.Text = Ps.ToString
        tb_portata.Text = qv.ToString
        tb_efficienza.Text = ni.ToString
        tb_rpm.Text = n.ToString
        lb_Pdiff.Text = Dp.ToString
        lb_Pcam1.Text = P_cam1.ToString
        lb_Pcam2.Text = P_cam2.ToString
        lb_Ta.Text = Ta.ToString
        lb_Hu.Text = Hu.ToString
        lb_Pa.Text = Pa.ToString
        lb_rho.Text = Ro1.ToString
        tb_tensione.Text = V_pr.ToString
        tb_corrente.Text = I_p.ToString
        tb_potenza.Text = We.ToString

    End Sub

    'CICLO DI ACQUISIZIONE

    Private Sub esegui_ciclo_di_acquisizione()

        leggi_misurazioni()

        esegui_calcoli()

        aggiorna_interfaccia_con_letture()

    End Sub

    'CLOCK DI LETTURA DATI E GESTIONE INSEGUIMENTI
    'DA VERIFICARE
    Private Sub Timer_prova_manuale_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_prova_manuale.Tick

        esegui_ciclo_di_acquisizione()

        If Stop_Max_Press Then
            passa_alla_fase_2()
        Else
            Select Case Fase
                Case 2 'punto massima pressione
                    acquisisci_max_press()

                Case 3 'termine test
                    Stop_prova()

                Case Else ' misura punti manuale

                    If (ch_portata.Checked) Then 'inseguimento portata
                        Insegui_Portata(qv, Qman, Q_Target_Ok)
                        ratio_raggiungimento_punto = Q_Target_Ok / num_cicli_lettura_stabile
                        If ratio_raggiungimento_punto > 1 Then ratio_raggiungimento_punto = 1
                        pb_raggiungimento_punto.Value = CType(ratio_raggiungimento_punto * 100, Integer)
                        riempi_buffer_circolari(Q_Target_Ok)
                        If (Q_Target_Ok > 0) Then 'abilito il pulsante salva punto solo se il valore è vicino al set point
                            bt_salva_punto.Enabled = True
                        Else
                            bt_salva_punto.Enabled = False
                        End If

                    ElseIf (ch_pressione.Checked) Then 'inseguimento pressione
                        Insegui_Pressione(Ps, Pman, Press_Target_Ok)
                        ratio_raggiungimento_punto = Press_Target_Ok / num_cicli_lettura_stabile
                        If ratio_raggiungimento_punto > 1 Then ratio_raggiungimento_punto = 1
                        pb_raggiungimento_punto.Value = CType(ratio_raggiungimento_punto * 100, Integer)
                        riempi_buffer_circolari(Press_Target_Ok)
                        If (Press_Target_Ok > 0) Then 'abilito il pulsante salva punto solo se il valore è vicino al set point
                            bt_salva_punto.Enabled = True
                        Else
                            bt_salva_punto.Enabled = False
                        End If

                    Else 'controllo manuale con Vaux
                        If (Vaux_percent = Vaux_percent_prev) And (serranda_percent = serranda_percent_prev) Then
                            Vaux_target_Ok = Vaux_target_Ok + 1
                        Else
                            If (Vaux_percent = Vaux_percent_prev) Then
                                Vaux_target_Ok = -30 ' forzo un attesa di 30 secondi prima di acquisire il dato dopo che ho cambiato il valore della serranda
                            Else
                                Vaux_target_Ok = -15 ' forzo un attesa di 15 secondi prima di acquisire il dato dopo che ho cambiato il valore di Vaux
                            End If
                        End If


                        ratio_raggiungimento_punto = Vaux_target_Ok / num_cicli_lettura_stabile
                        If ratio_raggiungimento_punto < 0 Then
                            ratio_raggiungimento_punto = 0
                        ElseIf ratio_raggiungimento_punto > 1 Then
                            ratio_raggiungimento_punto = 1
                        End If
                        pb_raggiungimento_punto.Value = CType(ratio_raggiungimento_punto * 100, Integer)
                        riempi_buffer_circolari(Vaux_target_Ok)
                        If (Vaux_target_Ok > 0) Then 'abilito il pulsante salva punto solo se il valore è vicino al set point
                            bt_salva_punto.Enabled = True
                        Else
                            bt_salva_punto.Enabled = False
                        End If

                    End If

                    serranda_percent_prev = serranda_percent
                    Vaux_percent_prev = Vaux_percent
                    aggiorna_uscita_Vaux()
                    aggiorna_uscita_serranda()

            End Select

        End If



    End Sub

    'PULSANTE DI SALVATAGGIO DATO

    Private Sub bt_salva_punto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_salva_punto.Click

        PopolaArray(nq)
        media_buffer_circolari() 'calcola i valori medi acquisiti
        esegui_calcoli() 'ricalcolo le grandezze da salvare

        MemorizzaDati(nq)
        nq = nq + 1

        'FINE PROVA PER ARRAY PIENO

        If nq > (num_max_punti - 1) Then Stop_prova()

    End Sub

    'DATI DI DEFAULT PARAMETRI AMBIENTALI
    'OK
    Sub imposta_dati_aria_default()

        Roref = 1.2
        Pa_calc = 101325
        Hu_calc_percent = 40.0
        Ta_calc = 20.0

        tb_rho_rif.Text = Roref
        tb_Ta_aria_rif.Text = Ta_calc
        tb_Hu_aria_rif.Text = Hu_calc_percent
        tb_Pb_aria_rif.Text = Pa_calc

    End Sub




    'IMPOSTA LE VARIABILI DI CONTROLLO INIZIALI
    'OK
    Public Sub inizializza_parametri_controllo()

        'questi 2 parametri non cambiano mai
        Tp = 6.5 '(numero tempi di campionamento)
        Tetap = Tcampionamento '(1 secondo)

        imposta_parametri_controllo_da_testo()
        'gli altri parametri li prendo inizialmente dalle caselle di testo
        'poi verranno cambiati a seconda della configurazione ugelli

    End Sub

    'INIZIALIZZAZIONE STRUMENTI
    'OK
    Public Function inizializza_strumenti() As Boolean

        'ritorna true se l'inizializzazione è andata a buon fine
        Dim esito As Boolean

        'acquisizione configurazione porte dai controlli del tab "impostazioni strumenti"
        nome_porta_P1 = cb_porta_P1.Text
        nome_porta_DP = cb_porta_DP.Text

        nome_porta_tachimetro = cb_porta_tachimetro.Text
        nome_porta_multimetro = cb_porta_multimetro.Text

        nome_porta_wattmetro = cb_porta_wattmetro.Text
        nome_porta_esam = cb_porta_esam.Text
        IPBarometro = tb_IP1.Text & "." & tb_IP2.Text & "." & tb_IP3.Text & "." & tb_IP4.Text

        'Ok_Analog, Ok_Barometro, Ok_Tacho, Ok_P1, Ok_DP, Ok_Wattmetro, Ok_Multimetro

        'inizializzazione connessioni strumenti

        Ok_Analog = aux_fan.inizializza()
        Ok_P1 = trasm_P1.inizializza(nome_porta_P1)
        Ok_DP = trasm_DP.inizializza(nome_porta_DP)

        Ok_Tacho = tachimetro.inizializza(nome_porta_tachimetro)
        Ok_Barometro = stazione_barometrica.inizializza(IPBarometro)

        If rb_infratek.Checked Then
            'se ho settato l'infratek inizializzo la porta Infratek
            Ok_Wattmetro = wattmetro.inizializza(nome_porta_wattmetro)
            bt_lettura_wattmetro.Enabled = True
            bt_lettura_esam.Enabled = False

        Else
            'altrimenti inizializzo la porta ESAM
            Ok_Wattmetro = esam.inizializza(nome_porta_esam)
            bt_lettura_wattmetro.Enabled = False
            bt_lettura_esam.Enabled = True
        End If


        If cb_frequenza_alimentazione.Text = "DC" Then
            Ok_Multimetro = multimetro.inizializza(nome_porta_multimetro)
            Ok_Multimetro = Ok_Multimetro AndAlso multimetro.imposta_lettura_IDC
            bt_lettura_multimetro.Enabled = True
        Else
            Ok_Multimetro = True
            ' Ok_Multimetro = multimetro.inizializza(nome_porta_multimetro)
        End If

        esito = Ok_Analog AndAlso Ok_P1 AndAlso Ok_DP AndAlso Ok_Wattmetro AndAlso Ok_Tacho AndAlso Ok_Multimetro AndAlso Ok_Barometro

        If esito Then
            lb_errore.Text = "NO ERROR"
            lb_errore.BackColor = Color.Lime
        Else
            lb_errore.Text = "ERRORE INIZIALIZZAZIONE STRUMENTI!!"
            lb_errore.BackColor = Color.Red
        End If

        Return esito

    End Function

    'TERMINAZIONE STRUMENTI
    'OK
    Public Function termina_strumenti() As Boolean

        'ritorna true se la terminazione è andata a buon fine
        Dim esito As Boolean

        Ok_Analog = aux_fan.termina()
        Ok_P1 = trasm_P1.termina()
        Ok_DP = trasm_DP.termina()

        Ok_Tacho = tachimetro.termina()
        Ok_Barometro = stazione_barometrica.termina()

        If rb_infratek.Checked Then
            Ok_Wattmetro = wattmetro.termina()
        Else
            Ok_Wattmetro = esam.termina()
        End If

        If cb_frequenza_alimentazione.Text = "DC" Then
            Ok_Multimetro = multimetro.termina()
            bt_lettura_multimetro.Enabled = False
        Else
            Ok_Multimetro = True
        End If

        esito = Ok_Analog AndAlso Ok_P1 AndAlso Ok_DP AndAlso Ok_Wattmetro AndAlso Ok_Tacho AndAlso Ok_Multimetro AndAlso Ok_Barometro

        If esito Then
            lb_errore.Text = "NO ERROR"
            lb_errore.BackColor = Color.Lime
        Else
            lb_errore.Text = "ERRORE TERMINAZIONE STRUMENTI!!"
            lb_errore.BackColor = Color.Red
        End If

        Return esito
    End Function

    'RIEMPIO LE STRUTTURE DATI RELATIVE AGLI UGELLI

    Private Function Qmax_ipotetica_data_conf_ugelli(ByVal num_conf_ugelli As UInt16) As UInt32

        Dim str = stringa_binaria_numconf(num_conf_ugelli) 'converto il numero di configurazione in una stringa di 0 e 1 corrispondente
        Dim portata_max As UInt32

        portata_max = 0
        For indice As Byte = 0 To (N_ugelli - 1)
            If str(indice) = "1" Then
                portata_max = portata_max + ugelli(N_ugelli - 1 - indice).Q_max
                'il primo carattere riguarda l'ultimo ugello
            End If
        Next

        Return portata_max

    End Function

    Private Sub inizializza_variabili_ugelli()

        'questi dati dipendono dalla configurazione fisica del cassone e dei sensori!!!!

        ugelli(0).diametro = 0.02
        ugelli(0).Q_max = 32
        ugelli(0).L_su_d = 0.6
        ugelli(0).k1_alfa = -7.006
        ugelli(0).k2_alfa = 134.6

        ugelli(1).diametro = 0.025
        ugelli(1).Q_max = 50
        ugelli(1).L_su_d = 0.6
        ugelli(1).k1_alfa = -7.006
        ugelli(1).k2_alfa = 134.6

        ugelli(2).diametro = 0.035
        ugelli(2).Q_max = 97
        ugelli(2).L_su_d = 0.6
        ugelli(2).k1_alfa = -7.006
        ugelli(2).k2_alfa = 134.6

        ugelli(3).diametro = 0.045
        ugelli(3).Q_max = 160
        ugelli(3).L_su_d = 0.5  'è l'unico con i parametri diversi per motivi costruttivi
        ugelli(3).k1_alfa = -6.688
        ugelli(3).k2_alfa = 131.5

        ugelli(4).diametro = 0.055
        ugelli(4).Q_max = 240
        ugelli(4).L_su_d = 0.6
        ugelli(4).k1_alfa = -7.006
        ugelli(4).k2_alfa = 134.6

        ugelli(5).diametro = 0.055
        ugelli(5).Q_max = 240
        ugelli(5).L_su_d = 0.6
        ugelli(5).k1_alfa = -7.006
        ugelli(5).k2_alfa = 134.6

        'calcolo la portata massima misurabile
        Qmax_assoluta = Qmax_ipotetica_data_conf_ugelli(N_configurazioni - 1)
        Qmax = 0

        'le configurazioni significative sono le seguenti:
        '1,2,3,4,6,7,10,11,13,15,22,25,27,29,31,56,59,63
        'riempio l'array delle configurazioni significative
        '(vedere file "parametri controllo.xls")

        configurazioni_significative(0) = 0
        configurazioni_significative(1) = 1
        configurazioni_significative(2) = 2
        configurazioni_significative(3) = 3
        configurazioni_significative(4) = 4
        configurazioni_significative(5) = 6
        configurazioni_significative(6) = 6
        configurazioni_significative(7) = 7

        configurazioni_significative(8) = 7
        configurazioni_significative(9) = 10
        configurazioni_significative(10) = 10
        configurazioni_significative(11) = 11
        configurazioni_significative(12) = 13
        configurazioni_significative(13) = 13
        configurazioni_significative(14) = 15
        configurazioni_significative(15) = 15

        configurazioni_significative(16) = 11
        configurazioni_significative(17) = 13
        configurazioni_significative(18) = 13
        configurazioni_significative(19) = 15
        configurazioni_significative(20) = 15
        configurazioni_significative(21) = 22
        configurazioni_significative(22) = 22
        configurazioni_significative(23) = 25

        configurazioni_significative(24) = 25
        configurazioni_significative(25) = 25
        configurazioni_significative(26) = 27
        configurazioni_significative(27) = 27
        configurazioni_significative(28) = 29
        configurazioni_significative(29) = 29
        configurazioni_significative(30) = 31
        configurazioni_significative(31) = 31

        configurazioni_significative(32) = 11
        configurazioni_significative(33) = 13
        configurazioni_significative(34) = 13
        configurazioni_significative(35) = 15
        configurazioni_significative(36) = 15
        configurazioni_significative(37) = 22
        configurazioni_significative(38) = 22
        configurazioni_significative(39) = 25

        configurazioni_significative(40) = 25
        configurazioni_significative(41) = 25
        configurazioni_significative(42) = 27
        configurazioni_significative(43) = 27
        configurazioni_significative(44) = 29
        configurazioni_significative(45) = 29
        configurazioni_significative(46) = 31
        configurazioni_significative(47) = 31

        configurazioni_significative(48) = 27
        configurazioni_significative(49) = 29
        configurazioni_significative(50) = 29
        configurazioni_significative(51) = 31
        configurazioni_significative(52) = 31
        configurazioni_significative(53) = 56
        configurazioni_significative(54) = 56
        configurazioni_significative(55) = 59

        configurazioni_significative(56) = 56
        configurazioni_significative(57) = 59
        configurazioni_significative(58) = 59
        configurazioni_significative(59) = 59
        configurazioni_significative(60) = 63
        configurazioni_significative(61) = 63
        configurazioni_significative(62) = 63
        configurazioni_significative(63) = 63

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    'PROCEDURA ESEGUITA ALL'APERTURA DEL PROGRAMMA
    'OK
    Private Sub Form1_shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown

        'INIZIALIZZAZIONE VARIABILI
        Larghezza = 0
        Altezza = 0
        inizializza_variabili_ugelli()

        inizializza_parametri_controllo()

        If inizializza_strumenti() Then

            TabControl1.SelectedIndex = 0 'mostro la scheda di impostazioni test

            aux_fan.azzera_uscita()
            aux_fan.abilita()

            imposta_dati_aria_default()

            'imposto nome file di default e percorso
            Path = "L:\Prove\" & System.DateTime.Today.Year.ToString
            NomeFile = System.DateTime.Today.Year.ToString & System.DateTime.Today.Month.ToString & System.DateTime.Today.Day.ToString
            tb_percorso_file_report.Text = Path & "\" & NomeFile & " report.xls"

            'imposto la variabile e il campo data
            Data = My.Computer.Clock.LocalTime.Date
            lb_data.Text = Data

            ch_vaux.Checked = True
            tb_portata.Text = 0.0

            '            Timer_generale.Interval = 50
            '            Timer_generale.Enabled = 1
            '            Timer_generale.Start()

        Else

            'se l'inizializzazione non ha avuto successo, mostro la scheda "impostazioni strumenti"
            TabControl1.SelectedIndex = 2

        End If

    End Sub

    'PROCEDURA ESEGUITA ALL'USCITA DAL PROGRAMMA
    'CHIUSURA PORTE E ARRESTO VENTILATORE AUSILIARIO

    Private Sub Form1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Leave

        aux_fan.azzera_uscita()
        termina_strumenti()

    End Sub

    Sub aggiorna_area_scarico()

        'aggiorno anche l'area di scarico dell'aspiratore
        Select Case cb_tipo_scarico.SelectedItem

            Case cb_tipo_scarico.Items(0)
                A2 = (Math.Pow(CDbl(tb_diametro_scarico.Text), 2) * 3.14159 / 4) / 1000000
            Case cb_tipo_scarico.Items(1)
                A2 = (tb_larghezza_scarico.Text * tb_altezza_scarico.Text) / 1000000
            Case cb_tipo_scarico.Items(2)
                A2 = tb_area_scarico.Text / 1000000
        End Select

    End Sub


    'PROCEDURE DI CAMBIO IMMAGINE UGELLO IN BASE ALLA SELEZIONE
    'OK
    Private Function tutti_ugelli_chiusi() As Boolean

        For indice = 0 To (N_ugelli - 1)
            If ugelli(indice).aperto Then
                tutti_ugelli_chiusi = False
                Exit Function
            End If
        Next
        tutti_ugelli_chiusi = True
    End Function

    Private Sub verifica_corretta_conf_ugelli()
        If tutti_ugelli_chiusi() Then
            lb_errore_configurazione.ForeColor = Color.Red
        Else
            aggiorna_parametri_da_configurazione_ugelli()
        End If
    End Sub

    Private Sub toggle_ugello_0()

        imposta_e_visualizza_stato_ugello_0(Not ugelli(0).aperto) 'inverto lo stato dell'ugello
        verifica_corretta_conf_ugelli()

    End Sub

    Private Sub imposta_e_visualizza_stato_ugello_0(ByRef aperto As Boolean)

        If aperto Then
            ugello_0.Image = immagine_ugello_aperto
            ugelli(0).aperto = True
            lb_errore_configurazione.ForeColor = Color.Transparent
        Else
            ugello_0.Image = immagine_ugello_chiuso
            ugelli(0).aperto = False
        End If

    End Sub

    Private Sub toggle_ugello_1()

        imposta_e_visualizza_stato_ugello_1(Not ugelli(1).aperto) 'inverto lo stato dell'ugello
        verifica_corretta_conf_ugelli()

    End Sub

    Private Sub imposta_e_visualizza_stato_ugello_1(ByRef aperto As Boolean)

        If aperto Then
            ugello_1.Image = immagine_ugello_aperto
            ugelli(1).aperto = True
            lb_errore_configurazione.ForeColor = Color.Transparent
        Else
            ugello_1.Image = immagine_ugello_chiuso
            ugelli(1).aperto = False
        End If

    End Sub

    Private Sub toggle_ugello_2()

        imposta_e_visualizza_stato_ugello_2(Not ugelli(2).aperto) 'inverto lo stato dell'ugello
        verifica_corretta_conf_ugelli()

    End Sub

    Private Sub imposta_e_visualizza_stato_ugello_2(ByRef aperto As Boolean)

        If aperto Then
            ugello_2.Image = immagine_ugello_aperto
            ugelli(2).aperto = True
            lb_errore_configurazione.ForeColor = Color.Transparent
        Else
            ugello_2.Image = immagine_ugello_chiuso
            ugelli(2).aperto = False
        End If

    End Sub

    Private Sub toggle_ugello_3()

        imposta_e_visualizza_stato_ugello_3(Not ugelli(3).aperto) 'inverto lo stato dell'ugello
        verifica_corretta_conf_ugelli()

    End Sub

    Private Sub imposta_e_visualizza_stato_ugello_3(ByRef aperto As Boolean)

        If aperto Then
            ugello_3.Image = immagine_ugello_aperto
            ugelli(3).aperto = True
            lb_errore_configurazione.ForeColor = Color.Transparent
        Else
            ugello_3.Image = immagine_ugello_chiuso
            ugelli(3).aperto = False
        End If

    End Sub

    Private Sub toggle_ugello_4()

        imposta_e_visualizza_stato_ugello_4(Not ugelli(4).aperto) 'inverto lo stato dell'ugello
        verifica_corretta_conf_ugelli()

    End Sub

    Private Sub imposta_e_visualizza_stato_ugello_4(ByRef aperto As Boolean)

        If aperto Then
            ugello_4.Image = immagine_ugello_aperto
            ugelli(4).aperto = True
            lb_errore_configurazione.ForeColor = Color.Transparent
        Else
            ugello_4.Image = immagine_ugello_chiuso
            ugelli(4).aperto = False
        End If

    End Sub

    Private Sub toggle_ugello_5()

        imposta_e_visualizza_stato_ugello_5(Not ugelli(5).aperto) 'inverto lo stato dell'ugello
        verifica_corretta_conf_ugelli()

    End Sub

    Private Sub imposta_e_visualizza_stato_ugello_5(ByRef aperto As Boolean)

        If aperto Then
            ugello_5.Image = immagine_ugello_aperto
            ugelli(5).aperto = True
            lb_errore_configurazione.ForeColor = Color.Transparent
        Else
            ugello_5.Image = immagine_ugello_chiuso
            ugelli(5).aperto = False
        End If

    End Sub

    Private Sub ugello_0_Click(sender As System.Object, e As System.EventArgs) Handles ugello_0.Click
        toggle_ugello_0()
    End Sub

    Private Sub Ugello_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugello_1.Click
        toggle_ugello_1()
    End Sub

    Private Sub Ugello_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugello_2.Click
        toggle_ugello_2()
    End Sub

    Private Sub Ugello_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugello_3.Click
        toggle_ugello_3()
    End Sub

    Private Sub Ugello_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugello_4.Click
        toggle_ugello_4()
    End Sub

    Private Sub ugello_5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugello_5.Click
        toggle_ugello_5()
    End Sub

    'procedure relative ai campi dimensionali del ventilatore
    'OK
    Private Sub aggiorna_controlli_tipo_scarico()

        Select Case cb_tipo_scarico.SelectedItem

            Case cb_tipo_scarico.Items(0) ' circolare
                tb_diametro_scarico.Enabled = True
                tb_larghezza_scarico.Enabled = False
                tb_altezza_scarico.Enabled = False
                tb_area_scarico.Enabled = False
                If (tb_diametro_scarico.Text = "") Then
                    lb_errore_diametro_scarico.ForeColor = Color.Red
                Else
                    lb_errore_diametro_scarico.ForeColor = Color.Transparent
                End If
                lb_errore_larghezza_scarico.ForeColor = Color.Transparent
                lb_errore_altezza_scarico.ForeColor = Color.Transparent
                lb_errore_area_scarico.ForeColor = Color.Transparent
                lb_errore_tipo_scarico.ForeColor = Color.Transparent

            Case cb_tipo_scarico.Items(1) 'rettangolare
                tb_larghezza_scarico.Enabled = True
                tb_altezza_scarico.Enabled = True
                tb_area_scarico.Enabled = False
                tb_diametro_scarico.Enabled = False
                If (tb_larghezza_scarico.Text = "") Then
                    lb_errore_larghezza_scarico.ForeColor = Color.Red
                Else
                    lb_errore_larghezza_scarico.ForeColor = Color.Transparent
                End If
                If (tb_altezza_scarico.Text = "") Then
                    lb_errore_altezza_scarico.ForeColor = Color.Red
                Else
                    lb_errore_altezza_scarico.ForeColor = Color.Transparent
                End If
                lb_errore_area_scarico.ForeColor = Color.Transparent
                lb_errore_diametro_scarico.ForeColor = Color.Transparent
                lb_errore_tipo_scarico.ForeColor = Color.Transparent

            Case cb_tipo_scarico.Items(2) 'altro
                tb_area_scarico.Enabled = True
                tb_larghezza_scarico.Enabled = False
                tb_altezza_scarico.Enabled = False
                tb_diametro_scarico.Enabled = False
                If (tb_area_scarico.Text = "") Then
                    lb_errore_area_scarico.ForeColor = Color.Red
                Else
                    lb_errore_area_scarico.ForeColor = Color.Transparent
                End If
                lb_errore_larghezza_scarico.ForeColor = Color.Transparent
                lb_errore_altezza_scarico.ForeColor = Color.Transparent
                lb_errore_diametro_scarico.ForeColor = Color.Transparent
                lb_errore_tipo_scarico.ForeColor = Color.Transparent
        End Select
    End Sub

    'SELEZIONE TIPOLOGIA DI SCARICO
    'OK
    Private Sub cb_tipo_scarico_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb_tipo_scarico.SelectedValueChanged
        aggiorna_controlli_tipo_scarico()
    End Sub

    'CALCOLO AREA RETTANGOLARE E DIAMETRO EQUIVALENTE
    'OK
    Private Sub tb_larghezza_scarico_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_larghezza_scarico.Leave

        Dim area, dimeq As Double

        If (tb_larghezza_scarico.Text <> "") Then Larghezza = tb_larghezza_scarico.Text
        area = Larghezza * Altezza
        dimeq = Math.Pow((area * 4 / 3.1415), 0.5)
        Virgola(dimeq, 1)
        tb_diametro_eq_scarico.Text = dimeq
        If (tb_larghezza_scarico.Text = "") Then
            lb_errore_larghezza_scarico.ForeColor = Color.Red
        Else
            lb_errore_larghezza_scarico.ForeColor = Color.Transparent
        End If

    End Sub
    'CALCOLO AREA RETTANGOLARE E DIAMETRO EQUIVALENTE
    'OK
    Private Sub tb_altezza_scarico_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_altezza_scarico.Leave

        Dim area, dimeq As Double

        If (tb_altezza_scarico.Text <> "") Then Altezza = tb_altezza_scarico.Text
        area = Larghezza * Altezza
        dimeq = Math.Pow((area * 4 / 3.1415), 0.5)
        Virgola(dimeq, 1)
        tb_diametro_eq_scarico.Text = dimeq
        If (tb_altezza_scarico.Text = "") Then
            lb_errore_altezza_scarico.ForeColor = Color.Red
        Else
            lb_errore_altezza_scarico.ForeColor = Color.Transparent
        End If

    End Sub
    'AREA DIVERSA (ALTRO) E CALCOLO DIAMETRO EQUIVALENTE
    'OK
    Private Sub Test_Area_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_area_scarico.Leave

        Dim diameq As Double

        If (tb_area_scarico.Text > 0) Then diameq = Math.Pow((tb_area_scarico.Text * 4 / 3.1415), 0.5)
        Virgola(diameq, 1)
        tb_diametro_eq_scarico.Text = diameq
        If (tb_area_scarico.Text = "") Then
            lb_errore_area_scarico.ForeColor = Color.Red
        Else
            lb_errore_area_scarico.ForeColor = Color.Transparent
        End If

    End Sub
    'VERIFICA INSERIMENTO DIAMETRO
    'OK
    Private Sub Test_Diametro_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_diametro_scarico.Leave

        If (tb_diametro_scarico.Text > 0) Then tb_diametro_eq_scarico.Text = tb_diametro_scarico.Text
        If (tb_diametro_scarico.Text = "") Then
            lb_errore_diametro_scarico.ForeColor = Color.Red
        Else
            lb_errore_diametro_scarico.ForeColor = Color.Transparent
        End If

    End Sub

    'verifica dati di configurazione ed eventuale avviso
    'OK
    Private Sub TabPage_test_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage_test.Enter

        If verifica_dati_impostazioni_test() Then

            Timer_dati_incompleti.Stop()

            aggiorna_area_scarico()

            aggiorna_parametri_da_configurazione_ugelli()

            If ((Vaux_percent = 0) And (Fase = 0)) Then 'se è la prima volta che entro devo impostare il timer idle
                Timer_idle.Interval = 1000  'intervallo 
                Timer_idle.Enabled = True
                Timer_idle.Start()
            End If

            If Not pressioni_azzerate Then
                Select Case MessageBox.Show("Vuoi azzerare i trasmettitori di pressione? (IL VENTILATORE DEV'ESSERE SPENTO!)", "Azzeramento sensori di pressione", MessageBoxButtons.YesNo)
                    Case DialogResult.Yes
                        azzera_trasmettitori_di_pressione()
                    Case DialogResult.No
                        pressioni_azzerate = True
                End Select
            End If

            aux_fan.abilita()
            aux_fan.abilita_uscita_serranda()

        Else

            Timer_dati_incompleti.Interval = 500 'avviso che mancano dei dati
            Timer_dati_incompleti.Enabled = True
            Timer_dati_incompleti.Start()

            TabControl1.SelectedIndex = 0 'ritorno sulla scheda impostazioni

        End If

    End Sub

    'FINESTRA DI AVVERTIMENTO DATI INCOMPLETI
    'OK
    Private Sub timer_dati_incompleti_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_dati_incompleti.Tick

        Popup_dati_incompleti.Show()
        Timer_dati_incompleti.Stop()

    End Sub

    'SISTEMA IN ATTESA DELL'AVVIAMENTO, LETTURA DATI a frequenza ridotta
    'OK
    'LETTURA DATI
    Private Sub timer_idle_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_idle.Tick

        esegui_ciclo_di_acquisizione()

    End Sub

    'ARRESTO LETTURA DATI SU USCITA DA PAGINA PRINCIPALE

    Private Sub TabPage_test_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage_test.Leave

        Timer_idle.Stop()

    End Sub

    'CARICAMENTO CONFIGURAZIONE DI DEFAULT DELLE PORTE SERIALI DEI SENSORI
    'OK
    Private Sub bt_carica_impostazioni_default_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_carica_impostazioni_default.Click

        CaricaImpostazioni(Camera_Aeraulica.My.Resources.impostazioni_strumenti_default)

    End Sub

    'SELEZIONE TIPO DI INSTALLAZIONE VENTILATORE IN PROVA
    'OK
    Private Sub cb_tipo_installazione_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_tipo_installazione.SelectedIndexChanged

        Install = cb_tipo_installazione.SelectedItem
        lb_errore_tipo_installazione.ForeColor = Color.Transparent

    End Sub

    'CARICAMENTO CONDIZIONI ATMOSFERICHE DI RIFERIMENTO DI DEFAULT
    'OK
    Private Sub bt_dati_aria_default_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_dati_aria_default.Click

        imposta_dati_aria_default()

    End Sub

    'ATTRIBUZIONE DEI VALORI DI RIFERIMENTO ATMOSFERICI ALLE VARIABILI DI PROGRAMMA
    'OK
    Private Sub tb_Ta_aria_rif_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_Ta_aria_rif.Leave

        If (tb_Ta_aria_rif.Text <> "") Then
            Ta_calc = tb_Ta_aria_rif.Text
        End If

    End Sub
    'OK
    Private Sub tb_Hu_aria_rif_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_Hu_aria_rif.Leave

        If (tb_Hu_aria_rif.Text <> "") Then
            Hu_calc_percent = CType(tb_Hu_aria_rif.Text, Single)
        End If

    End Sub
    'OK
    Private Sub tb_Pb_aria_rif_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_Pb_aria_rif.Leave

        If (tb_Pb_aria_rif.Text <> "") Then
            Pa_calc = tb_Pb_aria_rif.Text
        End If

    End Sub

    'PROCEDURE INNESCATE DA MODIFICA CHECKBOX CONTROLLI MANUALI
    'OK
    Private Sub ch_vaux_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_vaux.CheckedChanged

        aggiorna_controlli_inseguimento_vaux_e_serranda()

    End Sub
    'OK
    Private Sub ch_portata_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_portata.CheckedChanged

        aggiorna_controlli_inseguimento_portata()

    End Sub
    'OK
    Private Sub ch_pressione_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ch_pressione.CheckedChanged

        aggiorna_controlli_inseguimento_pressione()

    End Sub

    Public Sub azzera_trasmettitori_di_pressione()
        If (Not Timer_prova_automatica.Enabled) AndAlso (Not Timer_prova_manuale.Enabled) Then
            Timer_idle.Stop()
            trasm_P1.azzera()
            trasm_DP.azzera()
            Timer_idle.Start()
            pressioni_azzerate = True
        End If
    End Sub

    'AZZERAMENTO SENSORI DI PRESSIONE
    'OK
    Private Sub bt_azzera_pressioni_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_azzera_pressioni.Click

        azzera_trasmettitori_di_pressione()

    End Sub

    'REMINDER PRIMA DELL'AZZERAMENTO SENSORI DI PRESSIONE
    'OK
    Private Sub bt_azzera_pressioni_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_azzera_pressioni.MouseHover

        Status_prev = lb_messaggio.Text
        If (bt_azzera_pressioni.Enabled = True) Then lb_messaggio.Text = "VERIFICARE CHE TUTTI I MOTORI SIANO OFF"

    End Sub
    Private Sub bt_azzera_pressioni_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_azzera_pressioni.MouseLeave

        If (bt_azzera_pressioni.Enabled = True) Then lb_messaggio.Text = Status_prev

    End Sub

    'variazioni parametri in prova manuale

    'VARIAZIONE TENSIONE VENTILATORE AUSILIARIO PROVA MANUALE IN MANDATA

    Private Sub bt_inseguiVaux_mandata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_insegui_Vaux_mandata.Click
        Vaux_percent = tb_Vaux_target_mandata.Text
        aggiorna_uscita_Vaux()
    End Sub


    'VARIAZIONE TENSIONE VENTILATORE AUSILIARIO PROVA MANUALE IN ASPIRAZIONE

    Private Sub bt_insegui_Vaux_aspirazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_insegui_Vaux_aspirazione.Click
        Vaux_percent = tb_Vaux_target_aspirazione.Text
        aggiorna_uscita_Vaux()
    End Sub
    'VARIAZIONE PORTATA PROVA IN MANUALE

    Private Sub bt_insegui_portata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_insegui_portata.Click
        Try
            Qman = CType(tb_portata_target.Text, Double)
        Catch ex As Exception

        End Try

    End Sub
    'VARIAZIONE PRESSIONE PROVA IN MANUALE

    Private Sub bt_insegui_pressione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_insegui_pressione.Click
        Pman = tb_pressione_target.Text
    End Sub

    'CALCOLO DENSITA' DELL'ARIA DI RIFERIMENTO
    'OK
    Private Sub bt_calcola_rho_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_calcola_rho.Click

        Dim Psat_Ta_calc, Pv_calc, Teta_a_calc As Double

        If ((Ta_calc <> 0.0) And (Pa_calc > 0.0)) Then
            Psat_Ta_calc = 610.8 + 44.442 * Ta_calc + 1.4133 * Math.Pow(Ta_calc, 2) + 0.02768 * Math.Pow(Ta_calc, 3) + _
                0.000255667 * Math.Pow(Ta_calc, 4) + 0.00000289166 * Math.Pow(Ta_calc, 5)
            Pv_calc = Hu_calc_percent * Psat_Ta_calc / 100
            Teta_a_calc = Ta_calc + 273.15
            Roref = (Pa_calc - 0.378 * Pv_calc) / (287 * Teta_a_calc)
        End If
        Virgola(Roref, 3)
        tb_rho_rif.Text = Roref
    End Sub

    'GESTIONE IMPOSTAZIONE E VISUALIZZAZIONE CONFIGURAZIONI UGELLI
    Private Sub imposta_e_visualizza_configurazione_ugelli(ByVal num_conf_ugelli As UInt16)

        'ricavo la stringa che rappresenta lo stato di apertura ugelli
        Dim str As String = stringa_binaria_numconf(num_conf_ugelli)

        'i caratteri della stringa rappresentano lo stato degli ugelli
        imposta_e_visualizza_stato_ugello_5(str(0) = "1")
        imposta_e_visualizza_stato_ugello_4(str(1) = "1")
        imposta_e_visualizza_stato_ugello_3(str(2) = "1")
        imposta_e_visualizza_stato_ugello_2(str(3) = "1")
        imposta_e_visualizza_stato_ugello_1(str(4) = "1")
        imposta_e_visualizza_stato_ugello_0(str(5) = "1")

    End Sub

    'TROVO LA CONFIGURAZIONE UGELLI PROVA
    'OK
    Private Sub Genera_Config_Ug(ByVal Qrif As Double)

        Dim Qact As UInt32 'portata misurabile con una certa configurazione ugelli
        Dim Conf_Ug_Temp As UInt16

        'le configurazioni possibili sono molte, ma solo quelle significative vengono proposte
        'le configurazioni significative sono state prescelte in base alla portata misurabile
        ' (vedere sub "inizializza_variabili_ugelli")

        Conf_Ug_Temp = 0
        Do Until (Qact >= Qrif)
            Conf_Ug_Temp = Conf_Ug_Temp + 1
            If Conf_Ug_Temp > (N_configurazioni - 1) Then 'se sono arrivato oltre il massimo mi fermo
                Conf_Ug_Temp = N_configurazioni - 1
                Qact = Qmax_assoluta
                Exit Do
            End If
            Qact = Qmax_ipotetica_data_conf_ugelli(Conf_Ug_Temp)
        Loop

        'scelgo la configurazione significativa corrispondente a quella trovata dal programma
        Conf_Ug = configurazioni_significative(Conf_Ug_Temp)
        'applico la configurazione
        imposta_e_visualizza_configurazione_ugelli(Conf_Ug)


    End Sub

    'stimo la configurazione degli ugelli plausibile e la visualizzo

    Private Sub stima_configurazione_ugelli()

        If (tb_Qmax_rif.Text <> "") And IsNumeric(tb_Qmax_rif.Text) Then 'se ho scritto un numero nel campo portata
            If (CType(tb_Qmax_rif.Text, Double) > 0) Then 'se il numero è maggiore di 0
                Qmax = CType(tb_Qmax_rif.Text, Double)
                If Qmax > Qmax_assoluta Then
                    Qmax = Qmax_assoluta
                    tb_Qmax_rif.Text = Qmax.ToString
                End If
                Genera_Config_Ug(Qmax)
            End If
        End If


    End Sub


    'IMPOSTA CONFIGURAZIONE ADEGUATA IN BASE AL VALORE DI PORTATA INDICATO
    'OK
    Private Sub bt_genera_configurazione_ugelli_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_genera_configurazione_ugelli.Click

        stima_configurazione_ugelli()

    End Sub

    'OK
    Private Sub aggiorna_controlli_inseguimento_vaux_e_serranda()

        If (ch_vaux.Checked = True) Then
            Select Case cb_tipo_prova.Text
                Case "Mandata"
                    tb_Vaux_target_mandata.Visible = True
                    bt_insegui_Vaux_mandata.Visible = True
                    tb_Vaux_target_aspirazione.Visible = False
                    bt_insegui_Vaux_aspirazione.Visible = False
                Case "Aspirazione"
                    tb_Vaux_target_mandata.Visible = False
                    bt_insegui_Vaux_mandata.Visible = False
                    tb_Vaux_target_aspirazione.Visible = True
                    bt_insegui_Vaux_aspirazione.Visible = True
                Case Else
            End Select

            tb_percent_serranda_target.Visible = True
            bt_imposta_percent_serranda.Visible = True

            tb_portata_target.Visible = False
            bt_insegui_portata.Visible = False
            ch_portata.Checked = False

            tb_pressione_target.Visible = False
            bt_insegui_pressione.Visible = False
            ch_pressione.Checked = False
        End If

    End Sub
    'OK
    Private Sub aggiorna_controlli_inseguimento_portata()

        If (ch_portata.Checked = True) Then
            tb_portata_target.Visible = True
            bt_insegui_portata.Visible = True

            tb_Vaux_target_mandata.Visible = False
            bt_insegui_Vaux_mandata.Visible = False
            tb_Vaux_target_aspirazione.Visible = False
            bt_insegui_Vaux_aspirazione.Visible = False
            ch_vaux.Checked = False
            tb_percent_serranda_target.Visible = False
            bt_imposta_percent_serranda.Visible = False

            tb_pressione_target.Visible = False
            bt_insegui_pressione.Visible = False
            ch_pressione.Checked = False

        End If

    End Sub
    'OK
    Private Sub aggiorna_controlli_inseguimento_pressione()

        If (ch_pressione.Checked = True) Then
            tb_pressione_target.Visible = True
            bt_insegui_pressione.Visible = True

            tb_Vaux_target_mandata.Visible = False
            bt_insegui_Vaux_mandata.Visible = False
            tb_Vaux_target_aspirazione.Visible = False
            bt_insegui_Vaux_aspirazione.Visible = False
            ch_vaux.Checked = False
            tb_percent_serranda_target.Visible = False
            bt_imposta_percent_serranda.Visible = False

            tb_portata_target.Visible = False
            bt_insegui_portata.Visible = False
            ch_portata.Checked = False

        End If

    End Sub

    'VERIFICA SE L'UTENTE HA INSERITO I PERCORSI DI FOTO E REPORT
    'OK
    Private Function verifica_percorsi_file() As Boolean
        Dim errore As Boolean

        errore = ((tb_percorso_file_report.Text = "") Or (tb_percorso_file_foto.Text = ""))
        Return (Not errore)

    End Function

    'AVVIO PROVA IN AUTOMATICO
    'OK
    Private Sub bt_start_test_auto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_start_test_auto.Click

        If verifica_percorsi_file() Then
            Stop_Max_Press = False
            Test_Manuale = False
            reset_buffer_circolari()
            imposta_interfaccia_start_prova()
            imposta_interfaccia_prova_automatica()
            Start_Prova_Auto()
        Else
            MessageBox.Show("Verificare di aver inserito correttamente i percorsi di foto e report.")
        End If
    End Sub

    'AVVIO PROVA IN MANUALE
    'OK
    Private Sub bt_start_test_man_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_start_test_man.Click

        If verifica_percorsi_file() Then
            Stop_Max_Press = False
            Test_Manuale = True
            reset_buffer_circolari()
            imposta_interfaccia_start_prova()
            imposta_interfaccia_prova_manuale()
            Start_Prova_Man()
        Else
            MessageBox.Show("Verificare di aver inserito correttamente i percorsi di foto e report.")
        End If
    End Sub

    'IMPOSTAZIONE INTERFACCIA SE LA PROVA E' IN MANDATA O IN ASPIRAZIONE
    'OK
    Private Sub imposta_interfaccia_mandata_aspirazione()
        Select Case cb_tipo_prova.Text
            Case "Mandata"
                gb_Pcam1.Visible = True
                gb_Pcam2.Visible = False
                gb_Vaux_mandata.Visible = True
                gb_Vaux_aspirazione.Visible = False

            Case "Aspirazione"
                gb_Pcam1.Visible = False
                gb_Pcam2.Visible = True
                gb_Vaux_mandata.Visible = False
                gb_Vaux_aspirazione.Visible = True

            Case Else
        End Select
    End Sub

    'IMPOSTAZIONE INTERFACCIA E COMANDI PER START PROVA (SIA MANUALE CHE AUTOMATICA)
    'OK
    Private Sub imposta_interfaccia_start_prova()

        'disabilito i controlli non utilizzabili durante la prova

        'scheda test
        bt_azzera_pressioni.Enabled = False
        bt_dati_aria_default.Enabled = False
        bt_calcola_rho.Enabled = False
        tb_rho_rif.Enabled = False
        tb_Ta_aria_rif.Enabled = False
        tb_Hu_aria_rif.Enabled = False
        tb_Pb_aria_rif.Enabled = False

        tb_percorso_file_foto.Enabled = False
        tb_percorso_file_report.Enabled = False
        pb_percorso_file_foto.Enabled = False
        pb_percorso_file_report.Enabled = False

        bt_start_test_auto.Enabled = False
        bt_start_test_man.Enabled = False

        'abilito quelli utilizzabili solo durante la prova

        lb_raggiungimento_punto.Visible = True
        pb_raggiungimento_punto.Visible = True
        bt_fine_test.Enabled = True
        bt_max_press_termina_test.Enabled = True

        'scheda impostazioni test
        bt_genera_configurazione_ugelli.Enabled = False
        cb_tipo_installazione.Enabled = False
        cb_tipo_prova.Enabled = False

    End Sub

    'IMPOSTAZIONE INTERFACCIA E COMANDI PER PROVA AUTOMATICA
    'OK
    Private Sub imposta_interfaccia_prova_automatica()

        bt_salva_punto.Enabled = False

        ch_vaux.Enabled = False
        ch_portata.Enabled = False
        ch_pressione.Enabled = False

        bt_insegui_portata.Visible = False
        tb_portata_target.Visible = False
        bt_insegui_pressione.Visible = False
        tb_pressione_target.Visible = False
        bt_insegui_Vaux_aspirazione.Visible = False
        bt_insegui_Vaux_mandata.Visible = False
        tb_Vaux_target_aspirazione.Visible = False
        tb_Vaux_target_mandata.Visible = False
        bt_imposta_percent_serranda.Visible = False
        tb_percent_serranda_target.Visible = False


    End Sub

    'IMPOSTAZIONE INTERFACCIA E COMANDI PER PROVA MANUALE
    'OK
    Private Sub imposta_interfaccia_prova_manuale()

        bt_salva_punto.Enabled = True

        aggiorna_controlli_inseguimento_portata()
        aggiorna_controlli_inseguimento_pressione()
        aggiorna_controlli_inseguimento_vaux_e_serranda()
      
        If (ch_portata.Checked = True) Then
            tb_portata_target.Text = 1
            Qman = tb_portata_target.Text
        End If

        If (ch_pressione.Checked = True) Then
            tb_pressione_target.Text = 500
            Pman = tb_pressione_target.Text
        End If


    End Sub

    'REMINDER PRIMA DELLO START DELLA PROVA, ACCENDERE IL VENTILATORE!!!!
    'OK
    Private Sub bt_start_test_auto_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_start_test_auto.MouseHover

        Status_prev = lb_messaggio.Text
        If (bt_start_test_auto.Enabled = True) Then lb_messaggio.Text = "VERIFICARE CHE IL MOTORE IN TEST SIA ON"

    End Sub
    Private Sub bt_start_test_auto_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_start_test_auto.MouseLeave

        If (bt_start_test_auto.Enabled = True) Then lb_messaggio.Text = Status_prev

    End Sub

    'REMINDER PRIMA DELLO START DELLA PROVA, ACCENDERE IL VENTILATORE!!!!
    'OK
    Private Sub bt_start_test_man_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_start_test_man.MouseHover

        Status_prev = lb_messaggio.Text
        If (bt_start_test_man.Enabled = True) Then lb_messaggio.Text = "VERIFICARE CHE IL MOTORE IN TEST SIA ON"

    End Sub
    Private Sub bt_start_test_man_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_start_test_man.MouseLeave

        If (bt_start_test_man.Enabled = True) Then lb_messaggio.Text = Status_prev

    End Sub

    'ARRESTO PROVA

    Private Sub bt_fine_test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_fine_test.Click

        Fase = 3

    End Sub

    'IMPOSTAZIONE INTERFACCIA IN BASE A SELEZIONE TIPOLOGIA DI PROVA
    'OK
    Private Sub Test_Tipo_Prova_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb_tipo_prova.SelectedValueChanged

        imposta_interfaccia_mandata_aspirazione()

        lb_errore_tipo_prova.ForeColor = Color.Transparent

        Popup_tipo_prova.Show()

    End Sub

    Private Sub aggiorna_etichette_errori_impostazioni_test()
        'AGGIORNA LE ETICHETTE CHE EVIDENZIANO MANCANZE DI IMPOSTAZIONI TEST

        aggiorna_controlli_tipo_scarico()

        If (cb_tipo_installazione.Text = "") Then
            lb_errore_tipo_installazione.ForeColor = Color.Red
        Else
            lb_errore_tipo_installazione.ForeColor = Color.Transparent
        End If

        If (cb_tipo_prova.Text = "") Then
            lb_errore_tipo_prova.ForeColor = Color.Red
        Else
            lb_errore_tipo_prova.ForeColor = Color.Transparent
        End If

        If tutti_ugelli_chiusi() Then
            lb_errore_configurazione.ForeColor = Color.Red
        Else
            lb_errore_configurazione.ForeColor = Color.Transparent
        End If

        If tb_tensione_alimentazione.Text = "" Then
            lb_errore_tensione_alimentazione.ForeColor = Color.Red
        Else
            lb_errore_tensione_alimentazione.ForeColor = Color.Transparent
        End If

    End Sub

    'VERIFICA CHE TUTTI I DATI NECESSARI ALLA PROVA SIANO PRESENTI
    'OK
    Private Function verifica_dati_impostazioni_test() As Boolean

        Dim num_campi_errati As Integer

        num_campi_errati = 0
        If lb_errore_tipo_scarico.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_larghezza_scarico.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_altezza_scarico.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_area_scarico.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_diametro_scarico.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_tipo_installazione.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_tipo_prova.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_configurazione.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1
        If lb_errore_tensione_alimentazione.ForeColor = Color.Red Then num_campi_errati = num_campi_errati + 1

        If num_campi_errati = 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    'VERIFICA DELLE IMPOSTAZIONI MINIME PER L'ESECUZIONE DELLA PROVA
    'OK
    Private Sub TabPage_impostazioni_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage_impostazioni_test.Enter

        aggiorna_etichette_errori_impostazioni_test()
        aggiorna_controlli_tipo_scarico()

    End Sub


    'TEST STRUMENTI
    'OK
    Private Sub bt_lettura_barometro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_lettura_barometro.Click
        sw.Reset()
        sw.Start()

        stazione_barometrica.LeggiDati()
        tb_lettura_barometro.Text = stazione_barometrica.pressione.ToString + " " + _
                                    stazione_barometrica.temperatura.ToString + " " + _
                                    stazione_barometrica.umidità.ToString
        sw.Stop()
        tb_tempo_lettura_barometro.Text = sw.ElapsedMilliseconds.ToString

    End Sub

    Private Sub bt_lettura_P1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_lettura_P1.Click
        sw.Reset()
        sw.Start()

        tb_lettura_P1.Text = trasm_P1.pressione_Pa(Ok_P1).ToString
        sw.Stop()
        tb_tempo_lettura_P1.Text = sw.ElapsedMilliseconds.ToString
    End Sub

    Private Sub bt_lettura_DP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_lettura_DP.Click
        sw.Reset()
        sw.Start()

        tb_lettura_DP.Text = trasm_DP.pressione_Pa(Ok_DP).ToString
        sw.Stop()
        tb_tempo_lettura_DP.Text = sw.ElapsedMilliseconds.ToString
    End Sub

    Private Sub bt_lettura_wattmetro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_lettura_wattmetro.Click
        sw.Reset()
        sw.Start()

        tb_lettura_wattmetro.Text = wattmetro.tensione_V(Ok_Wattmetro).ToString & " " & wattmetro.corrente_A(Ok_Wattmetro).ToString & " " & wattmetro.potenza_W(Ok_Wattmetro).ToString

        sw.Stop()
        tb_tempo_lettura_wattmetro.Text = sw.ElapsedMilliseconds.ToString

    End Sub

    Private Sub bt_lettura_tachimetro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_lettura_tachimetro.Click
        sw.Reset()
        sw.Start()

        tb_lettura_tachimetro.Text = tachimetro.rpm.ToString

        sw.Stop()
        tb_tempo_lettura_tachimetro.Text = sw.ElapsedMilliseconds.ToString


    End Sub

    Private Sub bt_lettura_multimetro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_lettura_multimetro.Click
        sw.Reset()
        sw.Start()

        tb_lettura_multimetro.Text = multimetro.valore_letto(0).ToString

        sw.Stop()
        tb_tempo_lettura_multimetro.Text = sw.ElapsedMilliseconds.ToString


    End Sub

    Private Sub bt_lettura_esam_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_lettura_esam.Click
        sw.Reset()
        sw.Start()

        tb_lettura_esam.Text = esam.tensione(Ok_Wattmetro).ToString & " " & esam.corrente(Ok_Wattmetro).ToString & " " & esam.potenza(Ok_Wattmetro).ToString

        sw.Stop()
        tb_tempo_lettura_esam.Text = sw.ElapsedMilliseconds.ToString

    End Sub


    'REINIZIALIZZAZIONE STRUMENTI DOPO EVENTUALE CAMBIO CONFIGURAZIONE
    'OK
    Private Sub bt_reinizializza_strumenti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reinizializza_strumenti.Click

        termina_strumenti()

        inizializza_strumenti()

    End Sub

    Private Sub bt_imposta_parametri_controllo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_imposta_parametri_controllo.Click
        imposta_parametri_controllo_da_testo()
        aggiorna_parametri_controllo_effettivi()
    End Sub
    'verifica se i controlli del ventilatore ausiliario hanno un valore corretto, altrimenti lo corregge
    'OK
    Private Sub controlla_Vaux_mandata()
        If (CType(tb_Vaux_target_mandata.Text, Double) > Vaux_percent_max) Then tb_Vaux_target_mandata.Text = Vaux_percent_max.ToString
        If (CType(tb_Vaux_target_mandata.Text, Double) < Vaux_percent_min) Then tb_Vaux_target_mandata.Text = "0"
    End Sub
    'OK
    Private Sub controlla_Vaux_aspirazione()
        If (CType(tb_Vaux_target_aspirazione.Text, Double) > Vaux_percent_max) Then tb_Vaux_target_aspirazione.Text = Vaux_percent_max.ToString
        If (CType(tb_Vaux_target_aspirazione.Text, Double) < Vaux_percent_min) Then tb_Vaux_target_aspirazione.Text = "0"
    End Sub
    'OK
    Private Sub controlla_percent_serranda()
        If IsNumeric(tb_percent_serranda_target.Text) Then
            If (CType(tb_percent_serranda_target.Text, Double) > serranda_max) Then tb_percent_serranda_target.Text = serranda_max.ToString
            If (CType(tb_percent_serranda_target.Text, Double) < serranda_min) Then tb_percent_serranda_target.Text = serranda_min.ToString
        Else
            tb_percent_serranda_target.Text = ""
        End If

    End Sub
    'OK
    Private Sub tb_Vaux_target_mandata_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_Vaux_target_mandata.Leave
        controlla_Vaux_mandata()
    End Sub
    'OK
    Private Sub tb_Vaux_target_aspirazione_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_Vaux_target_aspirazione.Leave
        controlla_Vaux_aspirazione()
    End Sub
    'OK
    Private Sub tb_percent_serranda_target_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_percent_serranda_target.Leave
        controlla_percent_serranda()
    End Sub

    'OK
    Private Sub tb_portata_target_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_portata_target.Leave
        If (CType(tb_portata_target.Text, Double) > Qmax_assoluta) Then tb_portata_target.Text = Qmax_assoluta.ToString
        If (CType(tb_portata_target.Text, Double) < 0) Then tb_portata_target.Text = "0"
    End Sub
    'OK
    Private Sub tb_pressione_target_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_pressione_target.Leave
        If (CType(tb_pressione_target.Text, Double) > Pmax_assoluta) Then tb_pressione_target.Text = Pmax_assoluta.ToString
        If (CType(tb_pressione_target.Text, Double) < 0) Then tb_pressione_target.Text = "0"
    End Sub
    'OK
    Private Sub tb_num_punti_portata_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_num_punti_portata.Leave
        If CType(tb_num_punti_portata.Text, Integer) > num_max_punti_portata Then
            tb_num_punti_portata.Text = num_max_punti_portata.ToString
        ElseIf CType(tb_num_punti_portata.Text, Integer) < 1 Then
            tb_num_punti_portata.Text = "1"
        End If
    End Sub

    'APERTURA FINESTRA DI DIALOGO PER SELEZIONE FILE FOTO SETUP
    'OK
    Private Sub pb_percorso_file_foto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pb_percorso_file_foto.Click
        dialogo_percorso_foto.InitialDirectory = Path
        dialogo_percorso_foto.ShowDialog()
        If (dialogo_percorso_foto.FileName <> "") Then tb_percorso_file_foto.Text = dialogo_percorso_foto.FileName
    End Sub

    'APERTURA FINESTRA DI DIALOGO PER SELEZIONE FILE DI MEMORIZZAZIONE PROVA
    'OK
    Private Sub pb_nomefile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pb_percorso_file_report.Click

        dialogo_percorso_report.InitialDirectory = Path
        dialogo_percorso_report.ShowDialog()
        If (dialogo_percorso_report.FileName <> "") Then tb_percorso_file_report.Text = dialogo_percorso_report.FileName

    End Sub




    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'
    'PROCEDURE PROVA DI TENUTA
    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'

    Dim q_tenuta, q_tenuta_media, p_tenuta, p_tenuta_precedente, t_tenuta, t_tenuta_precedente, volume_tenuta As Double
    Dim buffer_letture_p(21), buffer_letture_t(21) As Double
    Dim riga As ULong
    Dim indice_buffer, indice_buffer_prec As Byte
    Dim cronometro1 As New Stopwatch

    Private Sub inizializza_report_tenuta()

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.Visible = True

        xlBook = xlApp.Workbooks.Open(tb_template_test_tenuta.Text, , True)
        xlSheet = xlBook.Worksheets("tabella")

        xlSheet.Cells(1, 1) = "t(s)"
        xlSheet.Cells(1, 2) = "P(Pa)"
        xlSheet.Cells(1, 3) = "Q(m3/h)"
        xlSheet.Cells(1, 4) = "Q media su 10s (m3/h)"

        xlBook.SaveAs(tb_percorso_report_test_tenuta.Text, , , , , , , Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution)

        cronometro1.Reset()
        cronometro1.Start()

        V = 0.2628      'volume di metà camera
        p_tenuta = 0
        t_tenuta = 0
        riga = 2
        indice_buffer = 0
    End Sub

    Private Sub chiudi_report_tenuta()

        cronometro1.Stop()

        xlBook.Save()
        xlBook.Close()
        xlApp.Quit()

        xlSheet = Nothing
        xlBook = Nothing
        xlApp = Nothing

    End Sub

    Private Sub scrivi_dato_tenuta()

        xlSheet.Cells(riga, 1) = t_tenuta
        xlSheet.Cells(riga, 2) = p_tenuta
        xlSheet.Cells(riga, 3) = q_tenuta
        If riga > 22 Then
            xlSheet.Cells(riga, 4) = q_tenuta_media
        End If

    End Sub

    Private Sub bt_inizia_prova_tenuta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_inizia_prova_tenuta.Click

        bt_inizia_prova_tenuta.Enabled = False
        bt_termina_prova_tenuta.Enabled = True

        inizializza_report_tenuta()

        Timer_prova_tenuta.Enabled = True
        Timer_prova_tenuta.Start()

    End Sub

    Private Sub bt_termina_prova_tenuta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_termina_prova_tenuta.Click

        bt_inizia_prova_tenuta.Enabled = True
        bt_termina_prova_tenuta.Enabled = False

        Timer_prova_tenuta.Stop()

        chiudi_report_tenuta()

    End Sub

    Private Sub Timer_prova_tenuta_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_prova_tenuta.Tick

        If indice_buffer = 20 Then
            indice_buffer_prec = 0
        Else
            indice_buffer_prec = indice_buffer + 1
        End If

        p_tenuta = trasm_P1.pressione_Pa(Ok_P1)
        buffer_letture_p(indice_buffer) = p_tenuta
        tb_p_tenuta.Text = p_tenuta.ToString

        t_tenuta = cronometro1.ElapsedMilliseconds / 1000
        buffer_letture_t(indice_buffer) = t_tenuta

        'calcolo portata
        q_tenuta = 3600 * V * (p_tenuta_precedente - p_tenuta) / (100000 * (t_tenuta - t_tenuta_precedente)) 'si ipotizza un pressione atmosferica di 100kPa
        Virgola(q_tenuta, 2)
        tb_q_tenuta_istantanea.Text = q_tenuta.ToString

        'calcolo portata 10s
        If riga > 22 Then
            q_tenuta_media = 3600 * V * (buffer_letture_p(indice_buffer_prec) - p_tenuta) / (100000 * (t_tenuta - buffer_letture_t(indice_buffer_prec)))
            Virgola(q_tenuta_media, 2)
            tb_q_tenuta_media.Text = q_tenuta_media
        End If

        scrivi_dato_tenuta()

        t_tenuta_precedente = t_tenuta
        p_tenuta_precedente = p_tenuta
        riga = riga + 1

        If indice_buffer = 20 Then
            indice_buffer = 0
        Else
            indice_buffer = indice_buffer + 1
        End If

    End Sub

    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'
    'PROCEDURE CARATTERIZZAZIONE SISTEMA
    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'
    '*****************************************************************************************************************************************'
    Dim t_caratterizzazione As Double


    '*******************************
    'CARATTERIZZAZIONE GUADAGNO
    '*******************************

    Private Sub inizializza_report_caratterizzazione()

        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlApp.Visible = True
        xlBook = xlApp.Workbooks.Add()
        xlSheet = xlBook.Worksheets("Foglio1")

        xlSheet.Cells(1, 1) = "Q(m3/h)"
        xlSheet.Cells(1, 2) = "DP(Pa)"
        xlSheet.Cells(1, 3) = "Ps(Pa)"
        xlSheet.Cells(1, 4) = "Vaux%"
        xlSheet.Cells(1, 5) = "Sum_Cj_d2"
        xlSheet.Cells(1, 6) = "Sum_d2"
        xlSheet.Cells(1, 7) = "t(s)"

        riga = 2

    End Sub

    Private Sub bt_inizia_caratterizzazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_inizia_caratterizzazione.Click
        bt_azzera_sensori_caratterizzazione.Enabled = False
        bt_inizia_caratterizzazione_dinamica.Enabled = False
        bt_termina_caratterizzazione_dinamica.Enabled = False
        bt_inizia_test_retroazione.Enabled = False

        bt_salva_punto_caratterizzazione.Enabled = True
        bt_termina_caratterizzazione.Enabled = True

        inizializza_report_caratterizzazione()

        cronometro1.Reset()
        cronometro1.Start()

        aux_fan.abilita()
        Timer_caratterizzazione_guadagno.Enabled = True
        Timer_caratterizzazione_guadagno.Start()

    End Sub

    Private Sub tabpage_caratterizzazione_sistema_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabpage_caratterizzazione_sistema.Enter

        'VERIFICA CHE TUTTI I DATI NECESSARI ALLA CARATTERIZZAZIONE SIANO PRESENTI

        If ((lb_errore_configurazione.ForeColor = Color.Red) Or (cb_tipo_prova.Text <> "Mandata")) Then
            TabPage_impostazioni_test.Select() ' ritorno sulle impostazioni
        Else
            aggiorna_parametri_da_configurazione_ugelli()
        End If

    End Sub

    Private Sub Timer_caratterizzazione_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_caratterizzazione_guadagno.Tick

        leggi_misure_caratterizzazione()
        calcoli_caratterizzazione()
        aggiorna_interfaccia_caratterizzazione()

    End Sub

    Private Sub tb_Vaux_percent_caratterizzazione_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_Vaux_percent_caratterizzazione.Leave
        Try
            If (CType(tb_Vaux_percent_caratterizzazione.Text, Double) > Vaux_percent_max) Then
                tb_Vaux_percent_caratterizzazione.Text = Vaux_percent_max.ToString
            ElseIf (CType(tb_Vaux_percent_caratterizzazione.Text, Double) < Vaux_percent_min) Then
                tb_Vaux_percent_caratterizzazione.Text = Vaux_percent_min.ToString
            ElseIf (CType(tb_Vaux_percent_caratterizzazione.Text, Double) > Vaux_percent_min) Then

            Else
                tb_Vaux_percent_caratterizzazione.Text = ""
            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub leggi_misure_caratterizzazione()

        'stazione barometrica
        Ok_Barometro = stazione_barometrica.LeggiDati()
        Ta = stazione_barometrica.temperatura
        Hu = stazione_barometrica.umidità
        Pa = stazione_barometrica.pressione
        T_ing_ug = Ta
        T_cam2 = Ta

        'pressioni
        Dp = trasm_DP.pressione_Pa(Ok_DP)
        If Ok_DP Then
            Virgola(Dp, 1)
            dp_prev = Dp
        Else
            Dp = dp_prev
        End If

        Ptemp = trasm_P1.pressione_Pa(Ok_P1)
        If Ok_P1 Then
            Virgola(Ptemp, 1)
            Ptemp_prev = Ptemp
        Else
            Ptemp = Ptemp_prev
        End If

        Select Case cb_tipo_prova.Text
            Case "Mandata"
                P_cam1 = Ptemp
                P_ing_ug = P_cam1
            Case "Aspirazione"
                P_cam2 = Ptemp
                P_ing_ug = P_cam2 + Dp
            Case Else
        End Select

        t_caratterizzazione = cronometro1.ElapsedMilliseconds / 1000

    End Sub

    Private Sub calcoli_caratterizzazione()

        Psat_Ta = 610.8 + 44.442 * Ta + 1.4133 * Math.Pow(Ta, 2) + 0.02768 * Math.Pow(Ta, 3) + 0.000255667 * Math.Pow(Ta, 4) + _
            0.00000289166 * Math.Pow(Ta, 5)
        Pv = Hu * Psat_Ta / 100
        Teta_a = Ta + 273.15
        Roa = (Pa - 0.378 * Pv) / (287 * Teta_a)
        Virgola(Roa, 3)
        Rw = Pa / (Teta_a * Roa)
        Ro1 = Roa

        P6 = Pa + P_ing_ug
        P4 = Pa + P_cam1
        Teta6 = T_ing_ug + 273.15
        Ro6 = P6 / (Teta6 * Rw)
        calcola_sum_cj_d2()
        qm = epsilon * 3.14159 * (Math.Sqrt(2 * Ro6 * Dp)) * Sum_Cj_d2

        'se c'è un errore di lettura mantengo il valore precedente
        If IsNumeric(qm) Then
            qm_prev = qm
        Else
            qm = qm_prev
        End If

        qv = 3600 * qm / Roref
        Virgola(qv, 1)
        Ps = P_cam1 * Roref / Ro6
        Virgola(Ps, 1)

    End Sub

    Private Sub aggiorna_interfaccia_caratterizzazione()
        tb_Ps_caratterizzazione.Text = Ps.ToString
        tb_qv_caratterizzazione.Text = qv.ToString
        tb_DP_caratterizzazione.Text = Dp.ToString
        tb_sum_cj_d2_caratterizzazione.Text = Sum_Cj_d2.ToString
        tb_sum_d2_caratterizzazione.Text = somma_d2_ugelli.ToString
        tb_t_caratterizzazione.Text = t_caratterizzazione.ToString
        tb_Vaux_percent_attuale.Text = Vaux_percent.ToString

    End Sub

    Private Sub bt_Vaux_percent_caratterizzazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_Vaux_percent_caratterizzazione.Click
        Try
            Vaux_percent = CType(tb_Vaux_percent_caratterizzazione.Text, Double)
            aggiorna_uscita_Vaux()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub salva_punto_caratterizzazione_su_report()

        xlSheet.Cells(riga, 1) = qv
        xlSheet.Cells(riga, 2) = Dp
        xlSheet.Cells(riga, 3) = Ps
        xlSheet.Cells(riga, 4) = Vaux_percent
        xlSheet.Cells(riga, 5) = Sum_Cj_d2
        xlSheet.Cells(riga, 6) = somma_d2_ugelli
        xlSheet.Cells(riga, 7) = t_caratterizzazione

        riga = riga + 1
    End Sub

    Private Sub bt_salva_punto_caratterizzazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_salva_punto_caratterizzazione.Click

        salva_punto_caratterizzazione_su_report()

    End Sub

    Private Sub bt_termina_caratterizzazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_termina_caratterizzazione.Click

        Timer_caratterizzazione_guadagno.Stop()
        cronometro1.Stop()

        aux_fan.azzera_uscita()

        bt_inizia_caratterizzazione.Enabled = True
        bt_azzera_sensori_caratterizzazione.Enabled = True
        bt_inizia_caratterizzazione_dinamica.Enabled = True
        bt_termina_caratterizzazione_dinamica.Enabled = True
        bt_inizia_test_retroazione.Enabled = True

        bt_salva_punto_caratterizzazione.Enabled = False
        bt_termina_caratterizzazione.Enabled = False

        xlBook.Save()

        xlSheet = Nothing
        xlBook = Nothing
        xlApp = Nothing

    End Sub

    Private Sub bt_azzera_sensori_caratterizzazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_azzera_sensori_caratterizzazione.Click
        trasm_P1.azzera()
        trasm_DP.azzera()
    End Sub

    '*******************************
    'CARATTERIZZAZIONE DINAMICA
    '*******************************

    Private Sub bt_inizia_caratterizzazione_dinamica_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_inizia_caratterizzazione_dinamica.Click

        bt_azzera_sensori_caratterizzazione.Enabled = False
        bt_inizia_caratterizzazione.Enabled = False
        bt_salva_punto_caratterizzazione.Enabled = False
        bt_termina_caratterizzazione.Enabled = False
        bt_inizia_test_retroazione.Enabled = False

        bt_inizia_caratterizzazione_dinamica.Enabled = False
        bt_termina_caratterizzazione_dinamica.Enabled = True

        inizializza_report_caratterizzazione()

        aux_fan.abilita()

        cronometro1.Reset()
        cronometro1.Start()

        Timer_caratterizzazione_dinamica.Enabled = True
        Timer_caratterizzazione_dinamica.Start()
    End Sub

    Private Sub Timer_caratterizzazione_dinamica_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_caratterizzazione_dinamica.Tick
        leggi_misure_caratterizzazione()
        calcoli_caratterizzazione()
        aggiorna_interfaccia_caratterizzazione()
        salva_punto_caratterizzazione_su_report()
    End Sub

    Private Sub bt_termina_caratterizzazione_dinamica_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_termina_caratterizzazione_dinamica.Click

        Timer_caratterizzazione_dinamica.Stop()
        cronometro1.Stop()

        aux_fan.azzera_uscita()


        bt_inizia_caratterizzazione.Enabled = True
        bt_azzera_sensori_caratterizzazione.Enabled = True
        bt_inizia_caratterizzazione_dinamica.Enabled = True
        bt_termina_caratterizzazione_dinamica.Enabled = True
        bt_inizia_test_retroazione.Enabled = True

        bt_salva_punto_caratterizzazione.Enabled = False
        bt_termina_caratterizzazione.Enabled = False

        xlBook.Save()

        xlSheet = Nothing
        xlBook = Nothing
        xlApp = Nothing

    End Sub

    '*******************************
    'TEST RETROAZIONE
    '*******************************
    Dim inseguimento_pressione As Boolean = False
    Dim inseguimento_portata As Boolean = False

    Private Sub tb_portata_target_caratterizzazione_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_portata_target_caratterizzazione.Leave
        Try
            If (CType(tb_portata_target_caratterizzazione.Text, Double) <= 0) Or (CType(tb_portata_target_caratterizzazione.Text, Double) > Qmax) Then
                tb_portata_target_caratterizzazione.Text = ""
            ElseIf CType(tb_portata_target_caratterizzazione.Text, Double) > 0 Then

            Else
                tb_portata_target_caratterizzazione.Text = ""
            End If

        Catch ex As Exception
            tb_portata_target_caratterizzazione.Text = ""
        End Try
    End Sub

    Private Sub tb_pressione_target_caratterizzazione_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_pressione_target_caratterizzazione.Leave
        Try
            If (CType(tb_pressione_target_caratterizzazione.Text, Double) <= 0) Or (CType(tb_pressione_target_caratterizzazione.Text, Double) > Qmax) Then
                tb_pressione_target_caratterizzazione.Text = ""
            ElseIf CType(tb_pressione_target_caratterizzazione.Text, Double) > 0 Then

            Else
                tb_pressione_target_caratterizzazione.Text = ""
            End If

        Catch ex As Exception
            tb_pressione_target_caratterizzazione.Text = ""
        End Try
    End Sub

    Private Sub bt_inizia_test_retroazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_inizia_test_retroazione.Click

        bt_azzera_sensori_caratterizzazione.Enabled = False
        bt_inizia_caratterizzazione_dinamica.Enabled = False
        bt_inizia_test_retroazione.Enabled = False
        bt_inizia_caratterizzazione.Enabled = False

        bt_termina_test_retroazione.Enabled = True

        cronometro1.Reset()
        cronometro1.Start()

        aux_fan.abilita()
        Vaux_percent = Vaux_percent_min
        aggiorna_uscita_Vaux()

        Timer_test_retroazione.Interval = Tcampionamento * 1000
        Timer_test_retroazione.Enabled = True
        Timer_test_retroazione.Start()
    End Sub

    Private Sub bt_termina_test_retroazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_termina_test_retroazione.Click

        Timer_test_retroazione.Stop()
        cronometro1.Stop()

        aux_fan.azzera_uscita()


        bt_azzera_sensori_caratterizzazione.Enabled = True
        bt_inizia_caratterizzazione_dinamica.Enabled = True
        bt_inizia_test_retroazione.Enabled = True
        bt_inizia_caratterizzazione.Enabled = True

        bt_termina_test_retroazione.Enabled = False

    End Sub

    Private Sub aggiorna_pulsanti_test_retroazione(ByVal inseguimento As Boolean)

        If inseguimento Then
            bt_insegui_portata_caratterizzazione.Enabled = False
            bt_insegui_pressione_caratterizzazione.Enabled = False
            bt_interrompi_retroazione.Enabled = True
        Else
            bt_insegui_portata_caratterizzazione.Enabled = True
            bt_insegui_pressione_caratterizzazione.Enabled = True
            bt_interrompi_retroazione.Enabled = False
        End If
    End Sub

    Private Sub Timer_test_retroazione_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_test_retroazione.Tick
        leggi_misure_caratterizzazione()
        calcoli_caratterizzazione()
        aggiorna_interfaccia_caratterizzazione()

        If inseguimento_portata Then
            Insegui_Portata(qv, Qman, Q_Target_Ok)

            If (Q_Target_Ok > num_cicli_lettura_stabile) Then
                inseguimento_portata = False
                Q_Target_Ok = 0
                aggiorna_pulsanti_test_retroazione(False)
            End If

            progress_caratterizzazione.Value = CType(100 * Q_Target_Ok / num_cicli_lettura_stabile, Integer)
        End If

        If inseguimento_pressione Then
            Insegui_Pressione(Ps, Pman, Press_Target_Ok)

            If (Press_Target_Ok > num_cicli_lettura_stabile) Then
                inseguimento_pressione = False
                Press_Target_Ok = 0
                aggiorna_pulsanti_test_retroazione(False)
            End If

            progress_caratterizzazione.Value = CType(100 * Press_Target_Ok / num_cicli_lettura_stabile, Integer)
        End If

        aggiorna_uscita_Vaux()


    End Sub

    Private Sub bt_insegui_portata_caratterizzazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_insegui_portata_caratterizzazione.Click
        Dim temp As Double

        Try
            temp = Qman
            Qman = CType(tb_portata_target_caratterizzazione.Text, Double)

            If Qman = 0 Then
                Qman = temp
            Else
                inseguimento_portata = True
                aggiorna_pulsanti_test_retroazione(True)
            End If


        Catch ex As Exception
            Qman = temp
        End Try

    End Sub

    Private Sub bt_insegui_pressione_caratterizzazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_insegui_pressione_caratterizzazione.Click
        Dim temp As Double

        Try
            temp = Pman
            Pman = CType(tb_pressione_target_caratterizzazione.Text, Double)

            If Pman = 0 Then tb_pressione_target_caratterizzazione.Text = "0"

            inseguimento_pressione = True
            aggiorna_pulsanti_test_retroazione(True)

        Catch ex As Exception
            Pman = temp
        End Try

    End Sub

    Private Sub bt_interrompi_retroazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_interrompi_retroazione.Click

        inseguimento_portata = False
        inseguimento_pressione = False
        aggiorna_pulsanti_test_retroazione(False)
        Vaux_percent = Vaux_percent_min
        aggiorna_uscita_Vaux()

    End Sub

    Private Sub cb_frequenza_alimentazione_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb_frequenza_alimentazione.SelectedValueChanged
        If cb_frequenza_alimentazione.Text = "DC" Then
            termina_strumenti()
            inizializza_strumenti()
        End If
    End Sub

    Private Sub tb_tensione_alimentazione_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tb_tensione_alimentazione.Leave

        If tb_tensione_alimentazione.Text = "" Then
            lb_errore_tensione_alimentazione.ForeColor = Color.Red
        Else
            Try
                V = CType(tb_tensione_alimentazione.Text, Double)
                lb_errore_tensione_alimentazione.ForeColor = Color.Transparent

            Catch ex As Exception
                lb_errore_tensione_alimentazione.ForeColor = Color.Red
            End Try
        End If


    End Sub

    Public Sub resetta_label_errori()
        lb_errore.BackColor = Color.Lime
        lb_errore.Text = "NO ERROR"
    End Sub

    'PROCEDURE PER SALVARE E RICHIAMARE IMPOSTAZIONI TEST
    'OK

    Private Sub pb_percorso_file_impostazioni_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pb_percorso_file_impostazioni.Click
        dialogo_percorso_file_impostazioni.InitialDirectory = Path
        dialogo_percorso_file_impostazioni.ShowDialog()
        If (dialogo_percorso_report.FileName <> "") Then tb_percorso_file_impostazioni_test.Text = dialogo_percorso_file_impostazioni.FileName
    End Sub

    Private Sub bt_salva_impostazioni_test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_salva_impostazioni_test.Click
        If tb_percorso_file_impostazioni_test.Text <> "" Then
            Try
                salva_impostazioni_test(tb_percorso_file_impostazioni_test.Text)
            Catch ex As Exception
                lb_errore_impostazioni.Text = ex.Message
            End Try

        End If


    End Sub

    Private Sub bt_carica_impostazioni_test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_carica_impostazioni_test.Click
        If tb_percorso_file_impostazioni_test.Text <> "" Then
            Try
                carica_impostazioni_test(tb_percorso_file_impostazioni_test.Text)
            Catch ex As Exception
                lb_errore_impostazioni.Text = ex.Message
            End Try

        End If
    End Sub


    Private Sub salva_impostazioni_test(ByRef nomefile As String)
        Dim tab As String = Chr(9)
        Dim stringa_dati As String = ""
        Dim str_conf_ug As String

        aggiorna_num_conf_ugelli()
        str_conf_ug = stringa_binaria_numconf(Conf_Ug) 'ugello aperto --> 1, ugello chiuso -->0

        stringa_dati = _
        str_conf_ug(2) & tab & _
        str_conf_ug(1) & tab & _
        tb_Qmax_rif.Text & tab & _
        tb_produttore_ventilatore.Text & tab & _
        tb_modello_ventilatore.Text & tab & _
        cb_tipo_scarico.Text & tab & _
        tb_tipo_ventilatore.Text & tab & _
        tb_larghezza_scarico.Text & tab & _
        tb_altezza_scarico.Text & tab & _
        tb_area_scarico.Text & tab & _
        tb_diametro_scarico.Text & tab & _
        tb_diametro_eq_scarico.Text & tab & _
        tb_note_ventilatore.Text & tab & _
        tb_produttore_motore.Text & tab & _
        tb_codice_motore.Text & tab & _
        tb_tensione_alimentazione.Text & tab & _
        cb_frequenza_alimentazione.Text & tab & _
        ch_correzione_tensione_alimentazione.Checked.ToString & tab & _
        cb_tipo_installazione.Text & tab & _
        cb_tipo_prova.Text & tab & _
        tb_esecutore.Text & tab & _
        tb_note_test.Text & tab & _
        tb_rif_prova.Text & tab & _
        tb_num_punti_portata.Text & tab & _
        ch_misura_portata_massima.Checked.ToString & tab & _
        str_conf_ug(3) & tab & _
        str_conf_ug(0)

        My.Computer.FileSystem.WriteAllText(nomefile, stringa_dati, False)
        lb_errore_impostazioni.Text = ""

    End Sub

    Private Sub carica_impostazioni_test(ByRef nomefile As String)
        Dim tab As String = Chr(9)
        Dim stringa_dati As String
        Dim array_dati() As String

        stringa_dati = My.Computer.FileSystem.ReadAllText(nomefile)
        array_dati = stringa_dati.Split(tab)

        imposta_e_visualizza_stato_ugello_1(CType(array_dati(0), Integer) = 1)
        imposta_e_visualizza_stato_ugello_2(CType(array_dati(1), Integer) = 1)
        tb_Qmax_rif.Text = array_dati(2)
        tb_produttore_ventilatore.Text = array_dati(3)
        tb_modello_ventilatore.Text = array_dati(4)
        cb_tipo_scarico.Text = array_dati(5)
        tb_tipo_ventilatore.Text = array_dati(6)
        tb_larghezza_scarico.Text = array_dati(7)
        tb_altezza_scarico.Text = array_dati(8)
        tb_area_scarico.Text = array_dati(9)
        tb_diametro_scarico.Text = array_dati(10)
        tb_diametro_eq_scarico.Text = array_dati(11)
        tb_note_ventilatore.Text = array_dati(12)
        tb_produttore_motore.Text = array_dati(13)
        tb_codice_motore.Text = array_dati(14)
        tb_tensione_alimentazione.Text = array_dati(15)
        cb_frequenza_alimentazione.Text = array_dati(16)
        ch_correzione_tensione_alimentazione.Checked = CType(array_dati(17), Boolean)
        cb_tipo_installazione.Text = array_dati(18)
        cb_tipo_prova.Text = array_dati(19)
        tb_esecutore.Text = array_dati(20)
        tb_note_test.Text = array_dati(21)
        tb_rif_prova.Text = array_dati(22)
        tb_num_punti_portata.Text = array_dati(23)
        ch_misura_portata_massima.Checked = CType(array_dati(24), Boolean)

        If array_dati.Length = 27 Then 'se è un file impostazioni nuovo imposto anche gli ugelli 2 e 3
            imposta_e_visualizza_stato_ugello_0(CType(array_dati(25), Integer) = 1)
            imposta_e_visualizza_stato_ugello_3(CType(array_dati(26), Integer) = 1)
        End If

        lb_errore_impostazioni.Text = ""

        aggiorna_etichette_errori_impostazioni_test()
    End Sub


    Private Sub rb_infratek_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_infratek.CheckedChanged
        If rb_infratek.Checked Then
            rb_esam.Checked = False
        Else
            rb_esam.Checked = True
        End If
    End Sub

    Private Sub rb_esam_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_esam.CheckedChanged
        If rb_esam.Checked Then
            rb_infratek.Checked = False
        Else
            rb_infratek.Checked = True
        End If
    End Sub

    Private Sub bt_imposta_percent_serranda_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_imposta_percent_serranda.Click

        If IsNumeric(tb_percent_serranda_target.Text) Then
            serranda_percent = Round(CType(tb_percent_serranda_target.Text, Double), 1)
            lb_percent_serranda.Text = serranda_percent.ToString
            aggiorna_uscita_serranda()
        End If

    End Sub

    Private Sub tb_percent_serranda_target_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_percent_serranda_target.TextChanged

    End Sub

    Private Sub tb_percorso_file_foto_TextChanged(sender As System.Object, e As System.EventArgs) Handles tb_percorso_file_foto.TextChanged

    End Sub
End Class