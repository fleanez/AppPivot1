Imports EEPHASE
Imports EEUTILITY.Enums

Public Class MyModel
    Implements IOpenModel

    Private Shared m_bObjectsCreated As Boolean = False
    Private Shared nMembership As Integer
    Private Shared nMyReportingPropertyEnum As Integer
    Private Shared strMyObjectName = "ObjectiveFunctionEq"
    Private Shared strMyReportProperty = "Objective Function"
    Private Shared strMySummaryReportProperty = "Objective Function"
    Private Shared strNewUnit = "$"
    Private Shared strNewSummaryUnit = "$"
    Private Shared nReportPhases As Integer() = New Integer(3) {SimulationPhaseEnum.LTPlan, SimulationPhaseEnum.PASA, SimulationPhaseEnum.MTSchedule, SimulationPhaseEnum.STSchedule}
    Private Shared strInputPath As String
    Private Shared strOutputPath As String

    Public Sub AfterInitialize() Implements IOpenModel.AfterInitialize

        'Initialize only once in MT Schedule:
        'Outside the loop to be used in batch execution too!
        If (G_oSTEP.IsFirstStep AndAlso G_oMODEL.SimulationPhase = SimulationPhaseEnum.MTSchedule) Then

            G_oMODEL.SolutionFile.AddReportingUnit(strNewUnit)
            G_oMODEL.SolutionFile.AddReportingUnit(strNewSummaryUnit)

            ' Add a REGION object to the solution dataset
            nMembership = G_oMODEL.SolutionFile.AddObject(ClassEnum.Region, strMyObjectName)

            ' Add a Reporting Property to the solution dataset
            nMyReportingPropertyEnum = G_oMODEL.SolutionFile.AddReportingProperty(CollectionEnum.SystemRegions, strMyReportProperty, strNewUnit, strSummaryName:=strMySummaryReportProperty, strSummaryUnit:=strNewSummaryUnit)

            ' Register a calculation function for MyReportProperty 
            G_oMODEL.SolutionFile.Register_CalculateIntervalData(CollectionEnum.SystemRegions,
                                                               nMembership,
                                                               nMyReportingPropertyEnum,
                                                               AddressOf MyCustomCalculation,
                                                               bWriteFlatFiles:=True,
                                                               nReportPhase:=nReportPhases,
                                                               bReportSamples:=True,
                                                               bReportStatistics:=True)

            G_oMODEL.SolutionFile.Register_CalculateSummaryData(CollectionEnum.SystemRegions,
                                                              nMembership,
                                                              nMyReportingPropertyEnum,
                                                              AddressOf MyCustomCalculation,
                                                              SummaryTypeEnum.Sum,
                                                              bWriteFlatFiles:=True,
                                                              nReportPhase:=nReportPhases)

            Dim strModelPath As String
            strModelPath = G_oMODEL.InputFilename
            strOutputPath = EEPHASE.G_oMODEL.OutputPath
            strInputPath = System.IO.Path.GetDirectoryName(strModelPath)

            Utils.CreateCSVfile(strOutputPath & "PivotOut.csv")

        End If

    End Sub

    Public Sub BeforeProperties() Implements IOpenModel.BeforeProperties
        'Throw New NotImplementedException()
    End Sub

    Public Sub AfterProperties() Implements IOpenModel.AfterProperties

        Dim nReserve As Integer
        Dim nGenerator As Integer
        Dim nPeriod As Integer
        Dim nPeriodsInStep As Integer
        Dim nIntervalsInStep As Integer
        Dim strModelName As String
        Dim lGeneratorReserve As New Dictionary(Of Integer, GeneratorReserve)
        Dim lGeneratorReserve_MaxSpareCapacity As New Dictionary(Of Integer, Double())
        Dim lGeneratorReserve_MaxRampCapacity As New Dictionary(Of Integer, Double())
        Dim lGeneratorReserve_MaxResponse As New Dictionary(Of Integer, Double())
        Dim lGeneratorReserve_ProvisionCoefficient As New Dictionary(Of Integer, Double())
        Dim lReserve_MinProvision As New Dictionary(Of String, Double())
        Dim dReserveCapacity As Double

        nIntervalsInStep = G_oSTEP.StepIntervalTo(G_oSTEP.CurrentStep) - G_oSTEP.StepIntervalFrom(G_oSTEP.CurrentStep) + 1
        nPeriodsInStep = G_oSTEP.LastPeriod - G_oSTEP.FirstPeriod + 1
        strModelName = G_oMODEL.Name

        Utils.OpenFile(strOutputPath & "PivotOut.csv")

        'Reserves min provisions (from reserve objects:
        For Each r As Reserve In ReservesIN
            If (r.IsInService) Then
                Dim dMinProvision(nPeriodsInStep) As Double
                For nPeriod = 1 To nPeriodsInStep
                    dMinProvision(nPeriod) = r(SystemReservesEnum.MinProvision, nPeriod)
                Next
                lReserve_MinProvision.Add(r.Name, dMinProvision)
            End If
        Next

        'Calculate the real and final Max Provision for each reserve and generator:
        For Each g As Generator In GeneratorsIN
            nReserve = 1
            Dim bUseMaxRampUp = g.IsDefined(SystemGeneratorsEnum.MaxRampUp)
            Dim bUseMaxRampDown = g.IsDefined(SystemGeneratorsEnum.MaxRampDown)

            For Each r As Reserve In g.Reserves

                'nMembership = G_oMODEL.SolutionFile.GetMembership(CollectionEnum.ReserveGenerators, r.Name, g.Name) 'Elegant but too inefficient
                nMembership = GeneratorReserve.GetID(g, r)
                lGeneratorReserve.Add(nMembership, New GeneratorReserve(nMembership, g, r))

                ' Register a calculation function for MyReportProperty (awesome idea! but it doesn't work!)
                'nMembership = G_oMODEL.SolutionFile.GetMembership(CollectionEnum.ReserveGenerators, r.Name, g.Name)
                'G_oMODEL.SolutionFile.Register_CalculateIntervalData(CollectionEnum.SystemGenerators,
                '                                                       nMembership,
                '                                                       nMyReportingPropertyEnum,
                '                                                       AddressOf Generator_ReserveMaxProvision,
                '                                                       bWriteFlatFiles:=True,
                '                                                       nReportPhase:=nReportPhases,
                '                                                       bReportSamples:=True,
                '                                                       bReportStatistics:=True)

                'G_oMODEL.SolutionFile.Register_CalculateSummaryData(CollectionEnum.SystemGenerators,
                '                                                      nMembership,
                '                                                      nMyReportingPropertyEnum,
                '                                                      AddressOf Generator_MaxProvision,
                '                                                      SummaryTypeEnum.Sum,
                '                                                      bWriteFlatFiles:=True,
                '                                                      nReportPhase:=nReportPhases)

                'CALCULATE MAX RESPONSE:
                Dim dMaxResponse(nPeriodsInStep) As Double
                Dim dRampMax(nPeriodsInStep) As Double
                Dim dSpare(nPeriodsInStep) As Double
                Dim dSpare2(nPeriodsInStep) As Double
                Dim dSpare3(nPeriodsInStep) As Double
                Dim dProvisionCoefficient(nPeriodsInStep) As Double

                nGenerator = r.Generators.IndexOf(g)

                Dim bUseMaxResponse As Boolean = r.Generators.IsDefined(nGenerator, EEUTILITY.Enums.ReserveGeneratorsEnum.MaxResponse)
                Dim bUseTimeFrame As Boolean = g.m_smReserves.ReserveIsRampLimited(nReserve) AndAlso r.IsDefined(SystemReservesEnum.Timeframe)

                For nPeriod = G_oSTEP.FirstPeriod To G_oSTEP.LastPeriod

                    dSpare(nPeriod) = g.RatedCapacity(nPeriod) - g.MinStableLevel(nPeriod)
                    dProvisionCoefficient(nPeriod) = 1.0 'This later will be fixed by constraints (for now just initialize in 1.0)
                    dRampMax(nPeriod) = dSpare(nPeriod)
                    If (bUseMaxRampUp AndAlso r.IsRaise) Then

                        If (bUseTimeFrame) Then
                            dRampMax(nPeriod) = 60 * r.MaxValue(EEUTILITY.Enums.SystemReservesEnum.Timeframe, nPeriod) * g.MaxValue(EEUTILITY.Enums.SystemGeneratorsEnum.MaxRampUp)
                        Else
                            dRampMax(nPeriod) = 60 * g(EEUTILITY.Enums.SystemGeneratorsEnum.MaxRampUp, nPeriod) * g.Units(nPeriod)
                        End If

                    ElseIf (bUseMaxRampDown AndAlso r.IsLower) Then

                        If (bUseTimeFrame) Then
                            dRampMax(nPeriod) = 60 * r.MaxValue(EEUTILITY.Enums.SystemReservesEnum.Timeframe, nPeriod) * g.MaxValue(EEUTILITY.Enums.SystemGeneratorsEnum.MaxRampDown)
                        Else
                            dRampMax(nPeriod) = 60 * g(EEUTILITY.Enums.SystemGeneratorsEnum.MaxRampDown, nPeriod) * g.Units(nPeriod)
                        End If

                    End If

                    If (bUseMaxResponse) Then
                        dMaxResponse(nPeriod) = g.Reserves_MaxResponse(nReserve, nPeriod)
                        If r.Type = EEUTILITY.Enums.ReserveTypeEnum.Operational Then
                            dMaxResponse(nPeriod) = Math.Max(dMaxResponse(nPeriod), g.Reserves_MaxReplacement(nReserve, nPeriod))
                        End If
                    Else
                        dMaxResponse(nPeriod) = -1.0 'too ugly??
                    End If

                Next

                lGeneratorReserve_MaxSpareCapacity.Add(nMembership, dSpare)
                lGeneratorReserve_MaxRampCapacity.Add(nMembership, dRampMax)
                lGeneratorReserve_MaxResponse.Add(nMembership, dMaxResponse)
                lGeneratorReserve_ProvisionCoefficient.Add(nMembership, dProvisionCoefficient)

                nReserve += 1
            Next
        Next

        'There are 3 fixes required: Non-simultaneous constraints, ajustment coefficients for min provision constraints and RHS as min provision (very ugly!)
        For Each c As Constraint In ConstraintsIN

            'We check only active constraints:
            If (c.IsInService) Then

                'Fix with the non-simultaneous contraints: (TODO: Move to a config file)
                If (c.Name.ToLower.Contains("_uniq")) Then
                    For nPeriod = 1 To nPeriodsInStep

                        Dim nLargestGenerator As Integer = 0
                        Dim dLargestCapacity As Double = 0.0

                        'Lets find out the largest generator
                        nGenerator = 1
                        For Each g As Generator In c.Generators
                            If (g.RatedCapacity(nPeriod) > dLargestCapacity) Then
                                nLargestGenerator = nGenerator
                                dLargestCapacity = g.RatedCapacity(nPeriod)
                            End If
                            nGenerator += 1
                        Next

                        'Now we set all others max provision to zero
                        nGenerator = 1
                        For Each g As Generator In c.Generators
                            If (nGenerator <> nLargestGenerator) Then
                                For Each r As Reserve In g.Reserves
                                    nMembership = GeneratorReserve.GetID(g, r)
                                    lGeneratorReserve_MaxSpareCapacity(nMembership)(nPeriod) = 0.0
                                    lGeneratorReserve_MaxRampCapacity(nMembership)(nPeriod) = 0.0
                                    lGeneratorReserve_MaxResponse(nMembership)(nPeriod) = 0.0
                                Next
                            End If
                            nGenerator += 1
                        Next
                    Next
                End If

                'Capture the reserve provision coefficient of constraints: (TODO: Move to config)
                If (c.Reserves.Count > 0 AndAlso c.Name.ToLower.Contains("minprovision")) Then
                    Dim r As Reserve
                    Dim dReserveProvisionCoefficient As Double
                    r = c.Reserves.First

                    'Only if one coefficient is defined, we need to keep record of it
                    If (c.Generators.IsDefined(ConstraintGeneratorsEnum.RaiseReserveProvisionCoefficient)) OrElse
                       (c.Generators.IsDefined(ConstraintGeneratorsEnum.LowerReserveProvisionCoefficient)) OrElse
                       (c.Generators.IsDefined(ConstraintGeneratorsEnum.RegulationRaiseReserveProvisionCoefficient)) OrElse
                       (c.Generators.IsDefined(ConstraintGeneratorsEnum.RegulationLowerReserveProvisionCoefficient)) OrElse
                       (c.Generators.IsDefined(ConstraintGeneratorsEnum.ReplacementReserveProvisionCoefficient)) Then

                        For nPeriod = 1 To nPeriodsInStep
                            nGenerator = 1
                            For Each g As Generator In c.Generators
                                nMembership = GeneratorReserve.GetID(g, r)
                                dReserveProvisionCoefficient = 1

                                'We have to use the workaround with MaxValue function (instead of by period values) because it's possible that it throws a nullreference exception
                                If (c.Generators.IsDefined(nGenerator, ConstraintGeneratorsEnum.RaiseReserveProvisionCoefficient)) Then
                                    'dReserveProvisionCoefficient = c.Generators(nGenerator, ConstraintGeneratorsEnum.RaiseReserveProvisionCoefficient, nPeriod)
                                    dReserveProvisionCoefficient = c.Generators.MaxValue(nGenerator, ConstraintGeneratorsEnum.RaiseReserveProvisionCoefficient)
                                ElseIf (c.Generators.IsDefined(ConstraintGeneratorsEnum.LowerReserveProvisionCoefficient)) Then
                                    'dReserveProvisionCoefficient = c.Generators(nGenerator, ConstraintGeneratorsEnum.LowerReserveProvisionCoefficient, nPeriod)
                                    dReserveProvisionCoefficient = c.Generators.MaxValue(nGenerator, ConstraintGeneratorsEnum.LowerReserveProvisionCoefficient)
                                ElseIf (c.Generators.IsDefined(nGenerator, ConstraintGeneratorsEnum.RegulationRaiseReserveProvisionCoefficient)) Then
                                    'dReserveProvisionCoefficient = c.Generators(nGenerator, ConstraintGeneratorsEnum.RegulationRaiseReserveProvisionCoefficient, nPeriod)
                                    dReserveProvisionCoefficient = c.Generators.MaxValue(nGenerator, ConstraintGeneratorsEnum.RegulationRaiseReserveProvisionCoefficient)
                                ElseIf (c.Generators.IsDefined(ConstraintGeneratorsEnum.RegulationLowerReserveProvisionCoefficient)) Then
                                    'dReserveProvisionCoefficient = c.Generators(nGenerator, ConstraintGeneratorsEnum.RegulationLowerReserveProvisionCoefficient, nPeriod)
                                    dReserveProvisionCoefficient = c.Generators.MaxValue(nGenerator, ConstraintGeneratorsEnum.RegulationLowerReserveProvisionCoefficient)
                                ElseIf (c.Generators.IsDefined(ConstraintGeneratorsEnum.ReplacementReserveProvisionCoefficient)) Then
                                    'dReserveProvisionCoefficient = c.Generators(nGenerator, ConstraintGeneratorsEnum.ReplacementReserveProvisionCoefficient, nPeriod)
                                    dReserveProvisionCoefficient = c.Generators.MaxValue(nGenerator, ConstraintGeneratorsEnum.ReplacementReserveProvisionCoefficient)
                                End If

                                If (lGeneratorReserve_ProvisionCoefficient.ContainsKey(nMembership)) Then 'We need to do this silly check since constraints may define coefficients for non-providers of that reserve
                                    lGeneratorReserve_ProvisionCoefficient(nMembership)(nPeriod) = dReserveProvisionCoefficient
                                End If
                                nGenerator += 1
                            Next
                            lReserve_MinProvision(r.Name)(nPeriod) = c(SystemConstraintsEnum.RHS, nPeriod)
                        Next
#If DEBUG Then
                        G_oFeedback.LogMessage("Modified min provision for reserve {0} from constraint {1} RHS", r.Name, c.Name)
#End If
                    End If

                End If

            End If

        Next

#If DEBUG Then
        'Write Generator - Reserve properties to file (interval data)
        Dim genres As GeneratorReserve
        For Each i As Integer In lGeneratorReserve_MaxSpareCapacity.Keys
            For nPeriod = 1 To nPeriodsInStep
                For nInterval = G_oSTEP.FirstIntervalInPeriod(nPeriod) To G_oSTEP.LastIntervalInPeriod(nPeriod)
                    genres = lGeneratorReserve(i)
                    For Each c As Company In genres.Generator.Companies

                        dReserveCapacity = Math.Min(lGeneratorReserve_MaxSpareCapacity(i)(nPeriod), lGeneratorReserve_MaxRampCapacity(i)(nPeriod))
                        If (lGeneratorReserve_MaxResponse(i)(nPeriod) <> -1.0) Then
                            dReserveCapacity = Math.Min(dReserveCapacity, lGeneratorReserve_MaxResponse(i)(nPeriod))
                        End If

                        Dim dMaxSpare As Double = lGeneratorReserve_MaxSpareCapacity(i)(nPeriod)
                        Dim dMaxRamp As Double = lGeneratorReserve_MaxRampCapacity(i)(nPeriod)
                        Dim dMaxResponse As Double = lGeneratorReserve_MaxResponse(i)(nPeriod)
                        Dim dProvisionCoefficient As Double = lGeneratorReserve_ProvisionCoefficient(i)(nPeriod)
                        Dim dMaxProvision As Double = dReserveCapacity * dProvisionCoefficient

                        Utils.AppendDataToCsv(strModelName, c.Name, genres.Reserve.Name, genres.Generator.Name, "ALL", nInterval, 1, dMaxSpare, dMaxRamp, dMaxResponse, dMaxProvision)
                    Next
                Next
            Next
        Next
#End If

        'Write totals per company:
        For Each c As Company In CompaniesIN
            Dim lCompanyMaxProvision As New Dictionary(Of Reserve, Double())
            For nPeriod = 1 To nPeriodsInStep
                For Each g As Generator In c.Generators
                    For Each r As Reserve In g.Reserves
                        nMembership = GeneratorReserve.GetID(g, r)
                        dReserveCapacity = Math.Min(lGeneratorReserve_MaxSpareCapacity(nMembership)(nPeriod), lGeneratorReserve_MaxRampCapacity(nMembership)(nPeriod))
                        If (lGeneratorReserve_MaxResponse(nMembership)(nPeriod) <> -1.0) Then
                            dReserveCapacity = Math.Min(dReserveCapacity, lGeneratorReserve_MaxResponse(nMembership)(nPeriod))
                        End If
                        Dim dProvisionCoefficient As Double = lGeneratorReserve_ProvisionCoefficient(nMembership)(nPeriod)
                        Dim dMaxProvision As Double = dReserveCapacity * dProvisionCoefficient
                        Dim dMaxProvisions(nPeriodsInStep) As Double
                        If (lCompanyMaxProvision.TryGetValue(r, dMaxProvisions)) Then
                            dMaxProvisions(nPeriod) += dMaxProvision
                        Else
                            ReDim dMaxProvisions(nPeriodsInStep)
                            dMaxProvisions(nPeriod) = dMaxProvision
                            lCompanyMaxProvision.Add(r, dMaxProvisions)
                        End If
                    Next
                Next
            Next
            For Each r As Reserve In lCompanyMaxProvision.Keys
                For nPeriod = 1 To nPeriodsInStep
                    For nInterval = G_oSTEP.FirstIntervalInPeriod(nPeriod) To G_oSTEP.LastIntervalInPeriod(nPeriod)
                        Utils.AppendDataToCsv(strModelName, c.Name, r.Name, "", "Max Provision", nInterval, 1, lCompanyMaxProvision(r)(nPeriod))
                    Next
                Next
            Next
        Next

        'DEPRECATED: This needs to be fixed by constraints. Ugly!
        'For Each r As Reserve In ReservesIN
        '    If (r.HasMinProvision) Then
        '        For nPeriod = 1 To nPeriodsInStep
        '            For nInterval = G_oSTEP.FirstIntervalInPeriod(nPeriod) To G_oSTEP.LastIntervalInPeriod(nPeriod)

        '                Utils.AppendDataToCsv(strModelName, "", G_oMODEL.SystemName, r.Name, nInterval, 1, dMinProvision, 0, 0, 0)

        '            Next
        '        Next
        '    End If
        'Next

        'Write Reserve properties (real Min provision) to file (interval data):
        For Each strReserve In lReserve_MinProvision.Keys
            For nPeriod = 1 To nPeriodsInStep
                For nInterval = G_oSTEP.FirstIntervalInPeriod(nPeriod) To G_oSTEP.LastIntervalInPeriod(nPeriod)
                    Utils.AppendDataToCsv(strModelName, "", G_oMODEL.SystemName, strReserve, "Min Provision", nInterval, 1, lReserve_MinProvision(strReserve)(nPeriod))
                Next
            Next
        Next

        Utils.CloseFile()

    End Sub

    Public Sub BeforeOptimize() Implements IOpenModel.BeforeOptimize
        'Throw New NotImplementedException()
    End Sub

    Public Sub AfterOptimize() Implements IOpenModel.AfterOptimize
        'Throw New NotImplementedException()
    End Sub

    Public Sub BeforeRecordSolution() Implements IOpenModel.BeforeRecordSolution

    End Sub

    Public Sub AfterRecordSolution() Implements IOpenModel.AfterRecordSolution

    End Sub

    Public Sub TerminatePhase() Implements IOpenModel.TerminatePhase
        'Throw New NotImplementedException()
    End Sub

    Public Function OnWarning(Message As String) As Boolean Implements IOpenModel.OnWarning
        'Throw New NotImplementedException()
    End Function

    Public Function EnforceMyConstraints() As Integer Implements IOpenModel.EnforceMyConstraints
        'Throw New NotImplementedException()
    End Function

    Public Function HasDynamicTransmissionConstraints() As Boolean Implements IOpenModel.HasDynamicTransmissionConstraints
        'Throw New NotImplementedException()
    End Function

    Public Function MyCustomCalculation() As Double()
        Dim m_values(Archive.StepPeriodCount) As Double
        m_values(0) = G_oMODEL.Task.ObjectiveValue()
        Return m_values
    End Function

End Class


Class GeneratorReserve

    Public Sub New(id As Integer, generator As Generator, reserve As Reserve)

        MembershipId = id
        Me.Generator = generator
        Me.Reserve = reserve

    End Sub

    Public Shared Function GetID(generator As Generator, reserve As Reserve) As Integer

        Return (reserve.Name & generator.Name).GetHashCode

    End Function

    Public Property Generator As Generator

    Public Property Reserve As Reserve

    Public Property MembershipId As Double

End Class


Class CompanyReservesIN

    Private Shared lCompanyMaxProvision As New Dictionary(Of Integer, Double())

    Public Shared Property MaxProvision(company As Company, reserve As Reserve, nPeriod As Integer) As Double
        Get
            Dim nMembership As Integer
            nMembership = CompanyReserve.GetID(company, reserve)
            Dim dMaxProvision(Archive.StepPeriodCount) As Double
            If (lCompanyMaxProvision.TryGetValue(nMembership, dMaxProvision)) Then
                Return dMaxProvision(nPeriod)
            Else
                Return 0.0
            End If
        End Get
        Set(value As Double)
            Dim nMembership As Integer
            nMembership = CompanyReserve.GetID(company, reserve)
            Dim dMaxProvision(Archive.StepPeriodCount) As Double
            If (lCompanyMaxProvision.TryGetValue(nMembership, dMaxProvision)) Then
                ReDim dMaxProvision(Archive.StepPeriodCount)
                dMaxProvision(nPeriod) = value
            Else
                dMaxProvision(nPeriod) = value
            End If
        End Set
    End Property

End Class

Class CompanyReserve

    Private MaxProvisions() As Double

    Public Sub New(company As Company, reserve As Reserve)
        Me.company = company
        Me.Reserve = reserve
        ReDim MaxProvisions(Archive.StepPeriodCount)
    End Sub

    Public Shared Function GetID(company As Company, reserve As Reserve) As Integer
        Return (reserve.Name & company.Name).GetHashCode
    End Function

    Public Property company As Company

    Public Property Reserve As Reserve

    Public Property MembershipId As Double

    Public Property MaxProvision(nPeriod As Integer) As Double
        Get
            Return MaxProvisions(nPeriod)
        End Get
        Set(value As Double)
            MaxProvisions(nPeriod) = value
        End Set
    End Property

End Class