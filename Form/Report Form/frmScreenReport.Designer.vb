<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScreenReport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.rptViewer = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.cmbDateType = New System.Windows.Forms.ComboBox()
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.lblEmployeeName = New System.Windows.Forms.Label()
        Me.cmbEmployeeName = New SergeUtils.EasyCompletionComboBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.btnClose = New PinkieControls.ButtonXP()
        Me.btnClear = New PinkieControls.ButtonXP()
        Me.btnGenerate = New PinkieControls.ButtonXP()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.lblEmploymentType = New System.Windows.Forms.Label()
        Me.cmbEmploymentType = New SergeUtils.EasyCompletionComboBox()
        Me.lblLeaveType = New System.Windows.Forms.Label()
        Me.cmbLeaveType = New SergeUtils.EasyCompletionComboBox()
        Me.pnlTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'rptViewer
        '
        Me.rptViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rptViewer.Location = New System.Drawing.Point(0, 120)
        Me.rptViewer.Name = "rptViewer"
        Me.rptViewer.Size = New System.Drawing.Size(1314, 481)
        Me.rptViewer.TabIndex = 2
        Me.rptViewer.TabStop = False
        '
        'cmbDateType
        '
        Me.cmbDateType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDateType.Font = New System.Drawing.Font("Verdana", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDateType.FormattingEnabled = True
        Me.cmbDateType.Items.AddRange(New Object() {"Leave Date", "Date Filed"})
        Me.cmbDateType.Location = New System.Drawing.Point(92, 5)
        Me.cmbDateType.Name = "cmbDateType"
        Me.cmbDateType.Size = New System.Drawing.Size(151, 24)
        Me.cmbDateType.TabIndex = 0
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.lblEmployeeName)
        Me.pnlTop.Controls.Add(Me.cmbEmployeeName)
        Me.pnlTop.Controls.Add(Me.lblStatus)
        Me.pnlTop.Controls.Add(Me.cmbStatus)
        Me.pnlTop.Controls.Add(Me.btnClose)
        Me.pnlTop.Controls.Add(Me.btnClear)
        Me.pnlTop.Controls.Add(Me.btnGenerate)
        Me.pnlTop.Controls.Add(Me.lblEndDate)
        Me.pnlTop.Controls.Add(Me.lblStartDate)
        Me.pnlTop.Controls.Add(Me.dtpEndDate)
        Me.pnlTop.Controls.Add(Me.dtpStartDate)
        Me.pnlTop.Controls.Add(Me.lblDate)
        Me.pnlTop.Controls.Add(Me.lblEmploymentType)
        Me.pnlTop.Controls.Add(Me.cmbEmploymentType)
        Me.pnlTop.Controls.Add(Me.lblLeaveType)
        Me.pnlTop.Controls.Add(Me.cmbLeaveType)
        Me.pnlTop.Controls.Add(Me.cmbDateType)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1314, 120)
        Me.pnlTop.TabIndex = 3
        '
        'lblEmployeeName
        '
        Me.lblEmployeeName.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmployeeName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEmployeeName.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblEmployeeName.ForeColor = System.Drawing.Color.Black
        Me.lblEmployeeName.Location = New System.Drawing.Point(6, 57)
        Me.lblEmployeeName.Name = "lblEmployeeName"
        Me.lblEmployeeName.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblEmployeeName.Size = New System.Drawing.Size(87, 24)
        Me.lblEmployeeName.TabIndex = 546
        Me.lblEmployeeName.Text = "Emp. Name"
        Me.lblEmployeeName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbEmployeeName
        '
        Me.cmbEmployeeName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbEmployeeName.Font = New System.Drawing.Font("Verdana", 9.5!)
        Me.cmbEmployeeName.FormattingEnabled = True
        Me.cmbEmployeeName.Items.AddRange(New Object() {"1503-001 RHODETTE R. EVANGELISTA", "1506-001 REYMOND M. BAYANI", "1506-002 RONALDO D. SY", "1510-001 EDEN L. CACA", "1510-002 RACHEL ANN M. CAJIPE", "1511-001 EDWIN S. MENDOZA", "1511-002 OLIVER A. MERIN", "1602-002 JONAMAY D. BENEDICTO", "1605-001 IMEE M. ARANAS", "1605-002 ANTHONY B. ATIENZA", "1606-001 GAYLORD A. BOHOL", "1607-001 IRISH D. LAJARA", "1701-001 AIDA D. ABESADA", "1701-003 CRISLYN V. AMANO", "1701-004 DESARIE V. ANDAL", "1701-005 SHELA-MAE M. APIL", "1701-007 ERMILYN B. OBNIALA", "1701-008 JEFFER B. BASAÑEZ", "1701-009 MYLENE M. BAULITA", "1701-014 DONNA O. CARAIG", "1701-016 JO-ANN MAY J. CARRILLO", "1701-020 MARY ROSE R. COMEL", "1701-022 JESTINE G. DE GUZMAN", "1701-024 ALELI APPLE R. DIAZ", "1701-025 ANIELYN D. VALENZUELA", "1701-065 CHRIS S. DIMAILIG", "1701-027 NESHILL A. DUMAGUING", "1701-028 ME-ANN A. ESPENILLA", "1701-029 SHIELA F. MALABUYOC", "1701-030 ANABELLE D. FRANCA", "1701-032 KAZEL M. GRIETA", "1701-033 NERMA V. GUZMAN", "1701-034 WENDY ANN M. HABLA", "1701-035 KRESIA ANNE C. HERNANDEZ", "1701-037 JULIE ANN A. JARITO", "1701-040 MA. NIÑA D. LIPAOPAO", "1701-041 MERESA M. LIWANAG", "1701-042 GENEDEL Q. MACINDO", "1701-043 LYKA S. MAGALLANES", "1701-044 MILDRED M. CASTILLO", "1701-045 YOLANDA O. MALACAMAN", "1701-047 MARY ROSE A. MARQUEZ", "1701-048 LANIE M. MALVEDA", "1701-052 MA. THERESA C. PEÑA", "1701-054 ALYSSA JANE S. PORCALLA", "1701-057 CHERRIE V. RECTO", "1701-058 JESSA A. REYES", "1701-059 INAH CHENTI D. RICOHERMOSO", "1701-062 LESLIE T. SALONGA", "1701-063 GINALYN R. SALOZA", "1701-066 EMMANUEL  JORDAN SANDOVAL", "1701-068 MYLENE D. TENORIO", "1701-069 KAYCEE S. UMALI", "1701-070 WILMA A. VALDEZ", "1701-071 DIANNE CAMILLE I. VILLANUEVA", "1701-075 MARY JANE B. MACAUBA", "1702-000 SATORU TOYOSHIMA", "1704-001 ELSA D. ACEDERA", "1704-006 LANNIE C. DELOS REYES", "1704-007 LILYBETH P. FAMODULAN", "1704-009 JAYVEE CLAIRE P. LOTERTE", "1704-011 JEMILYNE D. MATERUM", "1704-012 GENEROSE MAY P. OPEÑA", "1704-015 AIZA J. PROHIBIDO", "1704-017 SHIRLEY V. SAGUIT", "1704-018 REXCIE F. SANGALANG", "1704-019 RITCHELLE B. TAGLI", "1704-020 MARIAN A. LUMBRES", "1704-021 ARIANNE MEXIS FLOR P. ARRABE", "1704-022 MARIA KATHERINE B. BALAZON", "1704-023 JILLIAN B. ARANIEGO", "1704-024 GERLYN N. BIHIS", "1704-028 BUENA M. SERRANO", "1704-029 JOYLYN P. PETATE", "1704-030 MELISSA M. REYNA", "1704-032 CHARLENE L. VALENCIA", "1705-001 ANGELIQUE Z. ANTONA", "1705-002 RHINA ROSE V. CANTA", "1705-003 MAE JERAMY M. DE LA TORRE", "1705-006 DAISEREE E. VIERNES", "1705-007 KIMBERLY ANN G. BAROTILLA", "1705-008 RECHELLE B. GINES", "1705-009 HAZEL ANN G. ALVARIO", "1705-010 MARICAR J. PANGANIBAN", "1705-012 EDENROSE B. MAGISTRADO", "1705-013 ANGELIE O. MALIGALIG", "1705-016 WHYRINE C. MARANAN", "1705-019 LOTA M. LARA", "1705-021 VIENALEN B. ORPILLA", "1705-022 MARIA KAYE L. PEREZ", "1705-023 JHUNALYN R. PAGMANOJA", "1705-025 ALVIN J. ARAÑES", "1706-001 ESTEFHANNY E. LIM", "1706-002 MELODY S. CRUZAT", "1707-003 LOIDA B. LIBIRAN", "1707-004 ANGELYN I. ROXAS", "1708-001 CHEENEE M. DIAZ", "1708-002 MA. AILEEN F. AGUILA", "1708-003 FATIMA M. ATIENZA", "1708-004 ERICK M. BUNDALIAN", "1708-005 ARIEL C. GUEVARRA", "1709-001 MA. WELLA Q. CARBILLON", "1709-002 MYLEN A. IRINGAN", "1709-003 JOVIELYN P. MAGPANTAY", "1709-004 ARLENE S. ROXAS", "1709-006 CLARK ANTHONY B. VILLANUEVA", "1709-007 NICA C. VILLARIEZ", "1710-002 APRIL F. MORCOZO", "1710-003 MARIFE C. SAN JOSE", "1710-004 REZIEL C. ASUNCION", "1710-005 JESSAMIN C. CANDIDO", "1710-007 RICA D. CLEMENSO", "1710-010 CHARLENE M. EVANGELISTA", "1710-011 EDMAR D. FERRER", "1710-012 MARY ANN O. GUEVARRA", "1710-013 PAMELA M. HERNANDEZ", "1710-014 APRILYN T. JARIÑO", "1710-018 ZERLIE ANN T. PALAAC", "1710-020 JINKY T. SOLLER", "1710-021 CRISHELDA B. SUGANOB", "1711-002 JENNIFER S. BUNO", "1711-003 GERALDINE V. BUNSOL", "1711-005 MAILA M. EVANGELISTA", "1711-007 AIRENE B. FORBES", "1711-008 EDELIZA G. JAEN", "1711-010 RACHELLE H. RODRIGUEZ", "1711-011 JEAN M. ROMAN", "1801-002 VENGIE P. AGRIPA", "1801-003 HAZEL M. ALMAREZ", "1801-004 GLADYS L. ALMENDRAL", "1801-005 LIZA M. ALVAREZ", "1801-008 ARJELEN C. BARNEDO", "1801-009 CLARINE C. CALIVO", "1801-010 BABY JOY R. CARLOM", "1801-012 MARY JANE M. DESTA", "1801-013 GELLI E. FLORESCA", "1801-018 KAIRA MAE V. GARCIA", "1801-019 DANICA C. LASAT", "1801-021 NEMESIA H. MALABANAN", "1801-024 ROSELYN I. MANAY", "1801-026 ABEGAIL M. MEERA", "1801-028 DORRIS A. OBRADOR", "1801-029 RACHELL P. PILAPIL", "1801-030 IVY MAY L. PIÑON", "1801-031 ROSE-BELL B. SANGALANG", "1801-032 JASTINE A. TAPARANO", "1801-033 JIRLEEN M. TINA", "1802-001 CATHERINE L. DELA PEÑA", "1802-003 ALIZA V. BUENDIA", "1802-004 NELIZA G. CASTAÑARES", "1802-005 MICHELLE P. DELA CRUZ", "1802-006 ROSE B. DIMAYUGA", "1802-007 DANICA P. ILAG", "1802-008 MARIFE C. ILAO", "1802-009 ANGEL MAE U. INLOCNA", "1802-011 REA MAY M. LEYNES", "1802-014 MICHELL A. LOSA", "1802-015 RODELISA G. MARANAN", "1802-016 JULIE P. MILLA", "1802-019 MARY ANN S. ONAN", "1802-021 PRINCESS HAZEL G. VASQUEZ", "1803-001 AILA MAY A. FELESMINO", "1803-003 MA. CHRISTINE T. ALEGRIA", "1803-004 GERALDINE F. AQUINO", "1803-006 HAZEL CARALIPIO", "1803-007 JINKY D. CONSEBIDO", "1803-008 JOY O. FABELLA", "1803-009 RONEILYN G. GOLPEO", "1803-010 KAREN F. NORTE", "1803-011 RIHCA MEA C. SASOTA", "1803-016 ROWENA B. BRIONES", "1803-017 JULIET M. CASTILLO", "1803-019 LORENA J. JASARENO", "1803-020 JASPER C. JAVIER", "1803-021 RHAE MAE R. NAGUIMBING", "1803-022 RINA L. SOQUIAT", "1803-023 LIENCIE B. VELASCO", "1803-000 AKIO SHIMAMURA", "1805-002 CHARIE N. MOSCA", "1805-003 IRENE KRISTY S. TOREJAS", "1805-004 AIZA D. ROSALES", "1805-006 VINCENT ELOI A. MAGBOJOS", "1805-007 FLOREFE R. MALVEDA", "1805-008 MARICAR Q. MEERA", "1805-009 VANESSA E. QUIOZON", "1805-010 RACHELLE C. URGELLES", "1806-001 MERY R. DEDUYO", "1806-002 LANIE G. ESTAÑOL", "1806-003 JUDITH C. ORTEZ", "1806-005 SHERYL A. SICLON", "1806-006 LOVELYN A. TAYCO", "1807-002 HARRY J. TAÑEGA", "1807-003 LAARNIE L. ABANILLA", "1807-005 MARGIE A. BACULPO", "1807-006 MARY ANN L. BADASAN", "1807-007 JONALYN S. BALBA", "1807-008 ROSE MARIE A. DIENTRE", "1807-011 ELSA C. ESCALONA", "1807-012 GLAEZLE E. MANALO", "1807-015 JOAN L. ANDAYA", "1807-016 HAZEL E. MACASADIA", "1807-018 LEA R. MASUNCA", "1807-019 CLARIBEL H. MENDOZA", "1807-021 ERICA O. HABLA", "1807-022 ROCELYN B. OREA", "1807-023 VIA I. SANTOS", "1807-024 MA. DONA S. SEDUTAN", "1807-025 SHERLYN G. TUMAMBING", "1512-000 KENYA FUNAKI", "1809-001 NERISSA L. ANDES", "1809-002 MELANIE P. ARIOLA", "1809-003 DIVINA M. BALUDDA", "1809-004 MARJORIE T. BOLANTE", "1809-005 ARMANDO C. BUERE I", "1809-007 DIANE H. CASTILLO", "1809-008 NERISSA D. CATAPANG", "1809-009 ANGELICA B. CAYETANO", "1809-010 MARRY GRACE P. ESTEVES", "1809-012 MARY ANN S. IGNACIO", "1809-013 CRISTINA R. LEPALAM", "1809-015 DIANA JOY S. MADAMBA", "1809-016 JONATHAN R. MAGLUYAN", "1809-018 JACKIE LOU G. MALACAPO", "1809-020 GINALYN C. MANALO", "1809-022 ROWENA S. ORTEA", "1809-023 PERLA H. PACONIO", "1809-024 JASMINE D. PASOQUIN", "1809-026 CHERILYN M. ROBEL", "1809-027 INTISHAR M. SAGOSARA", "1809-028 MARY GRACE T. RIEGO", "1809-029 DANIEL BALVINO D. DALISAY", "1809-031 RICHARD M. PEÑALBA", "1809-030 VIRULYN P. GEMOTRA", "1809-032 PRECIL P. PRADO", "1810-001 REINA GLADYS M. AGUILA", "1810-003 SHELLA C. RECIO", "1810-007 SHENNA A. PADILLA", "1810-008 ERICSON S. VERGARA", "1811-002 JAY-AR H. AREGLADO", "1811-004 DINA T. BANDOJO", "1811-005 MARJORIE C. BAUTISTA", "1811-008 BRYAN ANTHONY B. BURCE", "1811-010 KAREN G. CAMPANA", "1811-011 GINALYN D. CANLOBO", "1811-014 ANGELICA G. COMIA", "1811-015 LENIE M. DELA PEÑA", "1811-018 GINA A. GARCIA", "1811-019 PRESCILA E. GONZALES", "1811-020 ZAIDA C. LOPEZ", "1811-021 ALEXA DANIELLE C. MANALO", "1811-022 SARAH D. MARQUEZ", "1811-023 JUVY B. PACRES", "1811-024 BRENEELYN O. PASOOT", "1811-025 MA. ELOISA D. RESURRECCION", "1811-026 CHENELYN L. SAJO", "1811-027 MARIA NIEVA G. SOFRANES", "1811-028 MARICAR M. TABAT", "1811-030 EDWIN S. SANCHEZ", "1811-031 JHON B. LONTOC", "1901-001 NOEMI D. DE VILLA", "1901-002 GLEN TRISTAN T. RICAMARA", "1901-005 LOVELY JOY D. ANAYA", "1901-006 RUTH P. ANCIADO", "1901-007 MARY GRACE L. BALURAN", "1901-009 SHARAJANE C. BRIZ", "1901-011 DYAN D. CANIETE", "1901-012 MARY JANE C. CUALQUERA", "1901-014 JAYBETH A. DIMAPALITAN", "1901-015 RIZELLE E. SACRISTIA", "1901-018 KAREN S. GONDRANEOS", "1901-019 MELBA M. MEJILLA", "1901-020 REBELYN C. MENDOZA", "1901-021 MARINETTE B. MONZON", "1901-025 MICHELLE ANN L. RAMOS", "1901-026 CARREN V. RAPSING", "1901-027 EFRELYN B. REPIA", "1901-028 RICHARD D. RIVERA", "1901-029 ELLA MAY I. SAGARIO", "1901-031 JUDY ANN C. UBALDO", "1901-032 LESLIE E. VISTA", "1901-033 NORIEL B. AQUINO", "1902-002 RUTCHEL O. ANDA", "1902-005 JENNIFER B. BELMONTE", "1902-006 MARY ANN BINLAYO", "1902-008 MARISTELA S. DE CASTRO", "1902-011 GEANETTE V. EBARDONE", "1902-012 MARIALYN ELARDO", "1902-014 ROSALIE M. MANGABAN", "1902-015 MANILYN A. MEDILLO", "1902-016 JOANNE L. OLA", "1902-017 BRENDA T. RIVERA", "1902-018 JOYFOLL E. RIVERA", "1902-019 JONALYN M. SILANG", "1902-020 REYNALDA V. TERO", "1902-021 JESSICA MARIE A. VILLANUEVA", "1902-023 ROSLYN H. VILLARBA", "1903-001 JOVELYN G. RENDORA", "1903-002 LAWRENCE B. ORINDAIN", "1903-003 ARIANE G. ARANTE", "1903-004 EMALYN D. ARMENTA", "1903-005 RONALYN P. AVETRIA", "1903-007 LYRA N. HULAR", "1903-008 GENNALYN T. ODOÑO", "1903-009 RUBY B. RELANES", "1903-011 CHRISTIAN M. NEDULA", "1903-012 MA. CRISTINA D. PEREZ", "1904-001 MARY GRACE M. ASENCE", "1904-002 MARY ANN B. LLAMERA", "1904-004 ANGELISA B. DE VERA", "1904-006 VISITACION E. GUMIRAN", "1904-007 MARY ROSE N. MAGSINO", "1904-010 JOAN R. MUSA", "1904-012 SHERLYN D. PAGUNSAN", "1904-015 CRESTIE G. SANTONIA", "1904-016 GEL D. SARNE", "1904-018 FRANCE JOYCE S. MANALO", "1905-001 MARRY ANN D. BATHAN", "1905-004 LORIE ANN U. CONSTANTINO", "1905-005 JOANNA E. DE LUNA", "1905-006 GEMMALYN O. DE TORRES", "1905-007 JOANNE V. DE LOS SANTOS", "1905-008 EMILY G. FLORES", "1905-010 PRECY M. MEDINA", "1905-011 APPLE M. MONTEJO", "1905-012 ANNA MARIE C. MORALLOS", "1905-015 JEYSA A. TARGA", "1905-016 MA. THEA R. TUGNAO", "1905-017 MA. CHRISELDA D. VILLEGAS", "1905-018 CARLA O. LUNA", "1906-002 CHARIE DELL S. CUADRO", "1906-003 ANABEL B. GALIT", "1906-004 JANELLE M. INDITA", "1906-005 RITCHEL ANN M. LUMANG", "1906-006 JUDY ANN MANGUSTA", "1906-009 ELLIE ROSE C. PALANG", "1906-010 RICHELLE M. REDOMA", "1906-011 DIANA LYKA R. RODIL", "1906-012 SANIKA MAE D. ROSAS", "1906-013 ROMALYN D. SOLOMON", "1907-002 DONNABEL V. CORNERO", "1907-003 MARIA CHERAMIE B. DOMINGO", "1907-004 REINA JOHAN R. GUTIERREZ", "1907-005 ANGELICA D. MADRIDANO", "1907-006 ROVELYN A. MARISTELA", "1907-007 ELSA A. PAGLINAWAN", "1907-008 KHEYNETH P. RODRIGO", "1907-009 MERILYN E. ROMAY", "1907-010 JESSICA V. SALVA", "1907-011 RECELYN E. SAN BUENAVENTURA", "1907-012 CHRISTINE MAY C. SERNA", "1907-013 NESTOR JR. T. TABIGNE", "1907-014 JANETH A. BELETARIO", "1907-015 NOVA S. CABALLERO", "1908-001 EVA L. ANTONES", "1908-002 ROSALIE C. ASI", "1908-003 JENNY ROSE M. ATIENZA", "1908-004 JUDITH N. AVILA", "1908-005 CHRISIEL M. BANTOC", "1908-006 JUDY ANN BAYLOSIS", "1908-007 JOYCE ANN D. CAMBAL", "1908-008 APOLONIA G. CUETO", "1908-009 BERNALYN G. DEQUITO", "1908-010 KIM LYNARD Y. HERNANDEZ", "1908-011 ROSE ANN C. HOLGADO", "1908-012 JESSICA T. LOZADA", "1908-013 REGINE JOY R. MUNDA", "1908-015 KATHLEEN G. PORRAS", "1908-017 CATHERINE A. ROBERTO", "1908-018 ANALYN D. SAUT", "1908-020 MARY GRACE E. SIONILLO", "1908-022 MAYBELYN P. TENORIO", "1908-023 RUTH O. YBARDOLAZA", "1908-025 MARY JANE M. PERALTA", "1908-024 MARICEL M. CATAPANG", "1909-001 MARENELL G. ALAYAN", "1909-002 BENILDA V. ARCIAGA", "1909-003 SHELLA M. BREDES", "1909-004 IRENE P. DELOS SANTOS", "1909-005 EDEZA M. DIZON", "1909-006 RODEL R. DOCTORA", "1909-007 MAITA S. LORILLA", "1909-008 AILYN N. MAGTIBAY", "1909-009 CHRIS ANN L. MAQUIÑANA", "1909-010 ANDREW C. NACARIO", "1909-011 JAN CAMILLE M. OJO", "1909-012 ARLENE C. ORLINA", "1909-014 JENNA MARIE S. PESIT", "1909-015 JUDYLYN M. ROMANO", "1909-016 RASHERL G. SENO", "1909-017 MICHAELA ROSE VIRTUCIO", "1909-018 JOBYLYN ZABALA", "1909-019 DAISY ANN D. DADULLA", "1910-003 MATETH N. ALISOSO", "1910-004 MELISA R. AMPER", "1910-005 MARICAR H. ARGAÑOSA", "1910-007 MELODY M. BAUTISTA", "1910-008 BERNADETTE B. BRIBON", "1910-010 CRISELDA T. CABANES", "1910-011 GERALDINE L. CAFE", "1910-012 APRIL G. CAMARA", "1910-013 MADELYN J. COLLADO", "1910-014 DYRINE A. DE CASTRO", "1910-018 MARK ANTHONY R. LIGAS", "1910-020 JENNY ROSE M. MANALO", "1910-021 YANIE N. NOYNAY", "1910-022 JOMIELYN I. TAN", "1910-023 JOSIE M. VERA", "1911-001 CHARITY C. BORJA", "1911-002 DECEREAN L. GARCIA", "1911-003 ANABEL P. LAURETA", "1911-004 SARAH JANE D. MAILOM", "1911-005 SALLIE M. MENDOZA", "1911-006 ROSELYN P. MIOLE", "1911-007 DIANNE M. OPEÑA", "1912-001 MARK RYAN R. CUYAPEN", "2001-007 IMELDA G. MARTINEZ", "2001-018 JOVI CLAIRE N. SANGALANG", "2001-005 RAMILYN N. DELA VEGA", "2001-016 JENNY G. BUTIAL", "2001-002 JESSA C. ATIENZA", "2001-006 JOLINA T. LANOZA", "2001-013 MIALYN C. TANDOY", "2001-010 LEIZA D. PACURIB", "2001-008 MARIEL D. MENDOZA", "2001-015 NELYN S. BALONDO", "2001-011 DAHLIA N. REDULA", "2001-009 MERY ROSE P. MONTERO", "2001-014 MICHA-ELLA A. TRIPULCA", "2001-017 CLAIRE B. DE LOS SANTOS", "2001-003 RALYN R. AGOJO", "2001-004 HAZZEL M. BAUTISTA", "2002-001 LESLIE E. MALABANAN", "2002-003 DONNA M. AGUILAR", "2002-004 OTCHE I. BALDOZ", "2002-005 JAIRO D. CABRAL", "2002-006 JESSA MAE R. CAMPO", "2002-007 GE-ANNE L. CANDELARIA", "2002-008 MANELYN D. DATINGUINOO", "2002-009 FRANCE JOYCE G. DIMACULANGAN", "2002-010 MARIBEL B. ESPLAGO", "2002-011 JANICE P. GARAY", "2002-012 JESSA D. HUBILLA", "2002-014 AMY L. MACASAYA", "2002-015 MARIA MERCEDES R. MAPANAO", "2002-016 AILEEN S. MONTECILLO", "2002-018 PRINCESS ODETTE NOGALES", "2002-019 KATHERINE S. OÑATE", "2002-020 AMOR H. PAHILANGCO", "2002-022 HAZEL ANN D. PIÑON", "2003-001 JOSEPH B. PUEBLO", "2004-002 GEMMA U. COROT", "2004-001 MARIAN A. BAGASINA", "2004-003 ANA MAE D. DELA GENTE", "2004-004 CHRISTIAN MAE H. DONAYRE", "2004-005 GRACE C. ESPIÑA", "2004-006 GIL R. LAPARAN", "2004-008 JESABEL P. MARTINEZ", "2004-009 MORIEL B. MENDOZA", "2004-011 JONEL P. OBAÑA", "2004-012 MAY MARIE B. OGORIDA", "2004-013 ARLENE P. PELAEZ", "2004-014 JESSICA M. PUNTANAR", "2004-016 CHELLO C. VELEZ", "2004-017 PRINCESS ROSE S. ZULUETA", "2005-003 REALEE M. ABANILLA", "2005-004 MARY GRACE L. AGUILAR", "2005-005 CHARMANE ROSE ANN D. ALAP", "2005-006 MERCY C. ALVAREZ", "2005-007 KRISTEL P. ATIENZA", "2005-008 SHIELA MAE E. AUSTRIA", "2005-009 MELANIE AZUELO", "2005-010 MARLYN M. BADILLO", "2005-011 GERALDENE V. BADO", "2005-012 EDELYN B. BALDE", "2005-013 LEZLIE C. BASCONCILLO", "2005-014 GERALDINE A. BIÑAS", "2005-015 CLAIRE R. CALIMOTAN", "2005-016 ANGELICA C. CASAPAO", "2005-017 LENY M. CASAPAO", "2005-018 SHERLYN P. CAYMO", "2005-019 MEAN G. DE LEON", "2005-020 LORENA L. DE LEON", "2005-021 RIZA L. DEL MUNDO", "2005-022 NIKKI M. DELOS SANTOS", "2005-023 CARMAY B. DILAO", "2005-024 CRISTINE M. DOMONDON", "2005-025 GLADYS L. GAMBOA", "2005-026 REGEN J. HOMERES", "2005-028 LALAINE J. MANCAY", "2005-029 STEFANIE MONTEVERDE", "2005-030 NIÑO TRISTAN H. NAMIA", "2005-032 EMERSON L. OLANDA", "2005-033 MA. ROSENDA G. ORTIZ", "2005-034 JOBELLE T. ORUGA", "2005-035 RHEA D. ORUGA", "2005-036 JAMIE G. PARSA", "2005-037 JONALYN T. PLATEROS", "2005-038 LERICKA JANE M. REDONDO", "2005-039 CLARISSA E. ROLL", "2005-040 MIA ROSE R. SALES", "2005-041 MA. TERESA M. SAN RAMON", "2005-043 JESSICA G. VALENZUELA", "2005-044 HEIDY O. VILLANUEVA", "2005-045 WENDY P. VILLANUEVA", "2005-031 GENEVA Z. NIMEDES", "2005-001 CHAD S. BELEN", "2005-002 ELDRIN BRYLLE M. LATAYAN", "2005-046 ANNEBELLE B. AYOP", "2005-047 ROBERT ANTHONY T. ALCANTARA", "2005-048 GLEZEL MAE E. GILO", "2005-049 ANGELICA M. NIEVA", "2006-001 JOYMEE Y. ABOT", "2006-002 FHERLYN E. AGUNOD", "2006-004 EVA MEY O. CASIL", "2006-005 MA. ALEXANDRA P. CASTILLO", "2006-006 BIANCA CAMILLE R. CATALUÑA", "2006-007 ROSEMARIE L. ESPERAGOZA", "2006-008 MARINEL G. GALAMGAM", "2006-009 JOVEL A. MALVEDA", "2006-010 ROSALYN L. MANJARES", "2006-011 MERRY CRIS B. MENDOZA", "2006-012 MILDRED E. MONTECILLO", "2006-013 MELODY C. REPOMANTA", "2006-015 COLENNE D. SUAREZ", "2006-016 SHEINA JOY D. TABLAR", "2006-017 JENNELYN N. TALARO", "2006-019 JAYSON M. TOLENTINO", "2006-020 CHRISTIAN M. TORRES", "2006-021 MEDELYN B. DELOS REYES", "2007-002 IVY M. CUETO", "2007-003 REYCHELLE B. DAVID", "2007-005 CINDY V. INCOMIO", "2007-006 MANILYN C. INOFRE", "2007-007 JESSALYN M. MAGPANTAY", "2007-008 JONALYN B. MOTILLA", "2008-001 LORNA B. BALOBAR", "2008-002 DAISY L. CAILAO", "2008-003 AIRALYN L. DIOLAS", "2008-005 DOROTHY D. PAGO", "2008-007 ARIEL Q. VILLAFLOR", "2009-002 EFRAIM JOSHUA B. MALIBONG", "2009-015 LIZA C. PASTRANA", "2009-003 MA. KEY KRISTINE E. ABELILLA", "2009-004 HANNIE MAE N. ALIVEN", "2009-005 MARY JOY M. CADAY", "2009-006 MARY GRACE A. DESEPEDA", "2009-007 MARY ANN C. DORADO", "2009-008 SHYRINE-ANN A. EREVILLA", "2009-010 JENNILYN H. HALUM", "2009-012 CHERRY H. LITERAL", "2009-013 JINKY C. MALIGO", "2009-014 MARIEFLOR M. MIRABALLES", "2010-003 RHONALYN P. CARANDANG", "2010-004 HYZEL M. CASTILLO", "2010-005 ESPERANZA A. CERVANTES", "2010-006 RUBY-ROSE ESTOR", "2010-007 MAYLEN L. ESTRADA", "2010-008 JOANA LESLIE M. ESTRELLA", "2010-009 ANALYN JEPOLLO", "2010-010 CORAZON Q. LEMUS", "2010-011 SHIENY C. MACARAIG", "2010-012 GLADYS G. MARCELO", "2010-013 MELDRID S. MASILANG", "2010-014 ZYLINE T. MENDOZA", "2010-015 JESSA A. MILITANTE", "2010-016 LEA T. PAMPLONA", "2010-017 ED CHRISTINE T. RUNAS", "2010-018 ALMA M. SABUITO", "2010-019 JULIET L. VILLANUEVA", "2011-001 JEVELYN R. ARGAÑOSA", "1907-000 YUKI YAMAMOTO", "2011-004 DOREEN MAY T. ANINGAT", "2011-005 MICHELLE S. ARCILLA", "2011-006 MAURICE E. ASIA", "2011-007 MA. ERICA B. BARJA", "2011-008 APRIL P. BELOLO", "2011-009 RYA PEARL S. CAGAT", "2011-010 MARY GRACE L. DEL PILAR", "2011-011 JOCELYN V. DELA CRUZ", "2011-012 CLEALYN A. FABRERO", "2011-013 IVY T. GERNALIN", "2011-014 MARY CRIS B. LITANA", "2011-015 FLORDELYN M. LUSTANIA", "2011-016 MAYRIN M. MACASADIA", "2011-017 LESLIE V. MAGBUJOS", "2011-018 DHENNY M. MANIPOL", "2011-019 JOAN L. MANONGSONG", "2011-020 JOCELYN A. MARQUEZ", "2011-021 REIGINE A. MENDOZA", "2011-023 GEMMA P. PALO", "2011-024 MADEL Y. PAMPLONA", "2011-026 ANGEL D. PAR", "2011-027 MARY JOY V. SEROMA", "2011-028 CAREN H. TORALBA", "2011-030 ROSE LYN M. VILLEGAS", "2101-001 AUGUSTINA F. ALERTA", "2101-002 JOAN B. BARRION", "2101-003 RENIEL J. BOTE", "2101-004 KIRK BRYAN B. CABAREÑO", "2101-006 EDMON C. DE GUZMAN", "2101-007 EDUARD JR. C. GUZMAN", "2101-008 JULIET R. LUGO", "2101-009 GERICO C. LUZ", "2101-010 RANDY E. NOGALES", "2101-011 ROSIE E. NOGALES", "2101-012 MICHELLE M. OSINSAO", "2101-013 EVERLYN H. CADANG", "2101-014 CECILE CENTENO", "2101-016 CAMILLE JOY D. DIMAANO", "2101-018 LORA F. LAYDEROS", "2101-019 ZENAIDA A. LLORCA", "2101-020 EVELYN S. LUMANGLAS", "2101-021 BABY SANILYN B. MALIPOL", "2101-022 GENELYN U. NITURAL", "2101-023 MARY ROSE A. NUÑEZ", "2101-024 JAYZELLE C. OLIVERA", "2101-025 CHERRY ANN A. RAMIREZ", "2101-026 MA. SHAIRA H. REY", "2101-027 CRISANTA A. RODELAS", "2101-028 ADONISA D. TAYCO", "2102-000 NORIKO NISHIDA", "2102-001 IRENE T. AALA", "2102-002 GLADYS A. ABEJUELA", "2102-003 BERNADETTE P. ADORNADO", "2102-004 ANALYN G. AMPUNIN", "2102-005 MARIA RICA H. ANDINO", "2102-006 ANGELICA E. ARIOLA", "2102-007 ALLYZA C. ARSENIAS", "2102-008 MARICAR A. ASUNCION", "2102-009 MELANIE R. AYCOCHO", "2102-010 MELBEVS V. BALEJONDA", "2102-011 JHONNALYN D. BAUTISTA", "2102-012 MARA M. BUSTAMANTE", "2102-013 RESEL MAE E. CARALIPIO", "2102-014 MARITES C. CASTILLO", "2102-015 MARY GRACE M. CLOSA", "2102-016 MARY GRACE B. CUDANIN", "2102-017 MARY ANN D. DELLOVA", "2102-018 MORRIEL ANN C. DIMEN", "2102-019 RUBIE V. DIOMAMPO", "2102-020 SAIDY L. DIZON", "2102-021 MARIA VANESSA R. DONGON", "2102-022 MA. GEMMA SARAH C. ENRIQUEZ", "2102-023 JUDY ANN EVALDEROZA", "2102-024 NECCA ANDREA R. FERNANDEZ", "2102-025 SHEILA S. GARCIA", "2102-026 MARJORIE B. GUERRERO", "2102-027 MYRA N. JAMISOLA", "2102-028 MARIBEL E. KALIBUSAN", "2102-029 LAYLINE A. LACOPIA", "2102-030 JEANCEL O. LANOZA", "2102-031 PRINCESS ANN A. LLAMES", "2102-032 JENALYN D. MANABAT", "2102-033 KAREN M. MARTINEZ", "2102-034 ROCHELLE MAUHAY", "2102-036 VERNADETH N. MONTALES", "2102-037 JOVELYN M. PEREZ", "2102-038 LORENA M. QUILESTE", "2102-039 JOONAMER T. ROBEJES", "2102-040 JUDY ANN G. SANCHA", "2102-041 ABEGAIL E. SANCHEZ", "2102-042 ROCHELL SANCHEZ", "2102-043 MICHELLE G. TANUYAN", "2102-044 ROXANNE M. VALENZUELA", "2102-046 ANDRIAN J. ZARA", "2103-000 HIROMI N. GAVILENO", "2103-001 JOSEPH M. REAMBONANZA", "2103-002 ODETTE P. ALCANTARA", "2103-003 MARIANNE A. ANDALLON", "2103-004 LORIEGENE B. CASTILLO", "2103-005 CHENEY MAY A. DELA PENA", "2103-007 VEVERLY V. ESTORES", "2103-008 MARY JANE T. HERNANDEZ", "2103-009 APPLE L. NACION", "2103-010 JENIE F. OBRADOR", "2103-011 IRICA D. PADILLA", "2103-012 RENILYN C. PASCUAL", "2103-013 LORRAINE G. PILI", "2103-014 ELLEN MAE P. PORRE", "2103-015 BABY JANE P. SABIGAN", "2103-016 LAARNI G. SAGUN", "2103-017 JOVELYN L. SALVACION", "2103-018 JESSICA B. SILABA", "2103-019 JOVY G. SOLAYBAR", "2103-020 CARMINA L. VILLANUEVA", "2104-003 MARIEANNE G. BARRION", "2104-012 NIÑA F. ABANID", "2104-013 NORIE B. ANDALIS", "2104-014 MARICON P. CASARINO", "2104-015 ROSE ANN L. CATINA", "2104-016 MA. CARINA CHARMAINE PIA", "2104-017 JANICE M. CRUZADO", "2104-018 CARREN C. DAÑO", "2104-019 MARY ACEL DERELO", "2104-021 KIM ABIGAIL V. GARCIA", "2104-022 MARY ANN P. GRAJO", "2104-023 ELOISA M. GREÑO", "2104-024 JAEMORE HERNANDEZ", "2104-025 SHEILA MARIE R. LORENZO", "2104-026 KHIMVERLY M. LUCERO", "2104-027 ARTH DOMINIQUE T. LUCILLO", "2104-028 MARY JANE S. MAASIN", "2104-029 JAYSON T. MARTINEZ", "2104-030 LESLIE ANN O. MERCADO", "2104-031 SHEENA MARIE C. MESIAS", "2104-032 KAYE ANN MAY L. MESOLANIA", "2104-033 EVELINDA L. OBSINADA", "2104-034 ZYRA MAY G. OJAS", "2104-035 LOREZA O. PACINOS", "2104-036 JERAH G. PANALIGAN", "2104-037 MARIMAR G. REJOSO", "2104-038 CATHERINE D. RID", "2104-039 CARLA M. ROSAS", "2104-040 EMERIE SABASCAR", "2104-041 MARY ROSE A. LINATOC", "2105-001 ROSHELL L. ARELLANO", "2105-002 BERNADETH C. BECHAYDA", "2105-003 JOY C. CALAGUAN", "2105-004 JERLYN C. CARINGAL", "2105-005 JAYNALYN B. SUNE", "2105-006 JENNELYN A. ZATARAIN", "2106-001 JOVELLY A. APUYAN", "2106-002 ANANEL A. BAYO", "2106-003 APRIL JANE T. BICO", "2106-004 GEMMALYN P. CARIÑO", "2106-005 JILL E. DIMAPILIS", "2106-006 MARVIN Q. MACATANGGA", "2106-007 JINGKIE L. MAGSINO", "2106-008 CAMILLE E. MALIPOL", "2106-009 ANGELICA C. MANIMTIM", "2106-010 HAYDEE A. MENDOZA", "2106-011 RANDYLEN P. MONDIGO", "2106-012 MIRASOL B. MORONG", "2106-013 ROSEMARIE U. SAN LORENZO", "2106-014 ANGELICA A. SARMIENTO", "2106-015 LALAINE P. SARMIENTO", "2106-016 AIZA W. TOLENTINO", "2106-017 MARK AERON F. CARANDANG", "2106-018 LEA MAE N. MANGUBAT", "2106-019 KYAN KHEN L. SY", "2106-020 MARY KARLIN JOYCE C. TANO", "2106-021 MACKY B. VILLANUEVA", "2106-023 KRIZHA SOL C. ANO", "2106-022 JAN VINCENT M. PEREZ", "2106-024 KRISTINE M. GARCIA", "2107-001 CATHERINE A. SIM", "2107-002 MELODY A. CONZON", "2107-003 ALELYN C. LUNA", "2107-004 MARY ANN A. MAOMAY", "2107-005 ROVELYN A. MENDOZA", "2107-006 REYMARIE L. OLIVERA", "2107-007 ANGEL E. ALINGATO ", "2107-008 REYNALYN C. BINGHOY", "2107-009 MADELINE M. GARCIA", "2107-010 KIMBERLY U. NATANAUAN", "2108-001 MA. LESLEY L. DE LUNA", "2108-002 SHYRA B. ESTOYA", "2108-003 ANTONINO T. GAVILENO", "21-001 SHANA ARGANDA", "21-002 ARA MAE R. BARRANDA", "21-003 RICA MAE C. MALAIBA", "21-004 JOHN KENNETH NAVARRO", "21-005 MARK JOSHUA P. PAZ", "21-006 KYLA MAE  T. RUBIO", "21-007 JOHN KENNETH O. RUBION", "21-008 MONICA V. ZATA", "2108-004 AIVY JOY A. CANTA", "2108-005 AILEEN L. CENTENO", "2108-006 MARRY JANE C. DE CHAVEZ", "2108-007 RIZZA G. ESMINO", "2108-008 BEVERLY F. FRONDA", "2108-009 LENDY M. GARGOLES", "2108-010 PAMELA A. LANGA", "2108-011 NECY T. LONTOC", "2108-012 JENNELYN C. MAGADIA", "2108-013 MYLENE D. MAGNO", "2108-014 NERISSA A. MALABANAN", "2108-015 MORENA C. MANALO", "2108-016 JOLINA D. PADUA", "2108-017 JIETH JOY D. PONCIANO", "2108-018 ANGELA P. REQUISO", "2108-019 JOSEPHINE IRENE ROJO", "2108-020 LIEZEL N. SABENIANO", "2108-021 JOY A. SONIEDO", "2108-022 CATHERENE V. TERCIAS", "2108-023 CRISTINE ELAINE A. VARELA", "2108-024 JOYME A. VILLANUEVA", "2108-025 MA. LOVELY P. ARIOLA", "2109-001 KATHERINE ANDAY", "2109-002 ALAYKA MAE A. ANGULO", "2109-003 JENNIFER S. CRUZ", "2109-004 KAYE ANN M. DORIA", "2109-005 DAISY H. FABRIQUEL", "2109-006 LADYLIN F. GABAY", "2109-007 JUDY ANN G. LAGUSTAN", "2109-008 MARY JOY R. MAYOR", "2109-009 DIVINA NOCOS", "2109-010 JANINE U. OLIVA", "2109-011 MARA JANE R. PARADERO", "2109-012 ALAISSA JOY G. PERMEJO", "2109-013 ROSALYN B. RAMOS", "2109-014 HIEZLE M. RIVERA", "2109-015 BEVERLY S. TURADO", "2109-016 MICHAEL D. UMALI", "2109-017 JEAN R. VARGAS", "2110-001 RENALYN T. RESABA", "2110-002 PRINCESS CAMILLE V. BAUTISTA", "2110-003 ERICA M. DOTE", "2110-004 EVALYN B. HAPIN", "2110-005 LARAH M. NAPAY", "2111-001 LIEZELLE P. MALDIA", "2111-002 MICHAEL R. BANABAN", "2111-003 JENNIFER D. BREIS"})
        Me.cmbEmployeeName.Location = New System.Drawing.Point(92, 57)
        Me.cmbEmployeeName.Name = "cmbEmployeeName"
        Me.cmbEmployeeName.Size = New System.Drawing.Size(1216, 24)
        Me.cmbEmployeeName.TabIndex = 545
        '
        'lblStatus
        '
        Me.lblStatus.BackColor = System.Drawing.SystemColors.Control
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblStatus.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblStatus.ForeColor = System.Drawing.Color.Black
        Me.lblStatus.Location = New System.Drawing.Point(586, 31)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblStatus.Size = New System.Drawing.Size(88, 24)
        Me.lblStatus.TabIndex = 544
        Me.lblStatus.Text = "Status"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbStatus
        '
        Me.cmbStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStatus.Font = New System.Drawing.Font("Verdana", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Items.AddRange(New Object() {"Leave Date", "Date Filed"})
        Me.cmbStatus.Location = New System.Drawing.Point(673, 31)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(635, 24)
        Me.cmbStatus.TabIndex = 5
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.btnClose.DefaultScheme = False
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnClose.Hint = "Close"
        Me.btnClose.Location = New System.Drawing.Point(1214, 85)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnClose.Size = New System.Drawing.Size(95, 30)
        Me.btnClose.TabIndex = 9
        Me.btnClose.TabStop = False
        Me.btnClose.Text = " Close"
        '
        'btnClear
        '
        Me.btnClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClear.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.btnClear.DefaultScheme = False
        Me.btnClear.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnClear.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnClear.Hint = "Clear filter"
        Me.btnClear.Image = Global.SickLeaveScreening.My.Resources.Resources.Undo_16_x_16
        Me.btnClear.Location = New System.Drawing.Point(1114, 85)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnClear.Size = New System.Drawing.Size(95, 30)
        Me.btnClear.TabIndex = 7
        Me.btnClear.TabStop = False
        Me.btnClear.Text = " Clear"
        '
        'btnGenerate
        '
        Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerate.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnGenerate.DefaultScheme = False
        Me.btnGenerate.DialogResult = System.Windows.Forms.DialogResult.None
        Me.btnGenerate.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.btnGenerate.Hint = "Generate report"
        Me.btnGenerate.Image = Global.SickLeaveScreening.My.Resources.Resources.OK_16_x_16
        Me.btnGenerate.Location = New System.Drawing.Point(1013, 85)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Scheme = PinkieControls.ButtonXP.Schemes.Blue
        Me.btnGenerate.Size = New System.Drawing.Size(95, 30)
        Me.btnGenerate.TabIndex = 6
        Me.btnGenerate.TabStop = False
        Me.btnGenerate.Text = "Generate"
        '
        'lblEndDate
        '
        Me.lblEndDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblEndDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEndDate.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblEndDate.ForeColor = System.Drawing.Color.Black
        Me.lblEndDate.Location = New System.Drawing.Point(416, 5)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblEndDate.Size = New System.Drawing.Size(48, 24)
        Me.lblEndDate.TabIndex = 543
        Me.lblEndDate.Text = "To"
        Me.lblEndDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblStartDate
        '
        Me.lblStartDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblStartDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblStartDate.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblStartDate.ForeColor = System.Drawing.Color.Black
        Me.lblStartDate.Location = New System.Drawing.Point(246, 5)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblStartDate.Size = New System.Drawing.Size(48, 24)
        Me.lblStartDate.TabIndex = 542
        Me.lblStartDate.Text = "From"
        Me.lblStartDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtpEndDate
        '
        Me.dtpEndDate.CustomFormat = ""
        Me.dtpEndDate.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEndDate.Location = New System.Drawing.Point(463, 5)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(120, 24)
        Me.dtpEndDate.TabIndex = 2
        '
        'dtpStartDate
        '
        Me.dtpStartDate.CustomFormat = ""
        Me.dtpStartDate.Font = New System.Drawing.Font("Verdana", 10.0!)
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStartDate.Location = New System.Drawing.Point(293, 5)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(120, 24)
        Me.dtpStartDate.TabIndex = 1
        '
        'lblDate
        '
        Me.lblDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDate.Font = New System.Drawing.Font("Verdana", 8.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate.ForeColor = System.Drawing.Color.Black
        Me.lblDate.Location = New System.Drawing.Point(6, 5)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblDate.Size = New System.Drawing.Size(87, 24)
        Me.lblDate.TabIndex = 383
        Me.lblDate.Text = "Date"
        Me.lblDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEmploymentType
        '
        Me.lblEmploymentType.BackColor = System.Drawing.SystemColors.Control
        Me.lblEmploymentType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEmploymentType.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblEmploymentType.ForeColor = System.Drawing.Color.Black
        Me.lblEmploymentType.Location = New System.Drawing.Point(6, 31)
        Me.lblEmploymentType.Name = "lblEmploymentType"
        Me.lblEmploymentType.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblEmploymentType.Size = New System.Drawing.Size(87, 24)
        Me.lblEmploymentType.TabIndex = 538
        Me.lblEmploymentType.Text = "Emp. Type"
        Me.lblEmploymentType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbEmploymentType
        '
        Me.cmbEmploymentType.Font = New System.Drawing.Font("Verdana", 9.5!)
        Me.cmbEmploymentType.FormattingEnabled = True
        Me.cmbEmploymentType.Items.AddRange(New Object() {"1503-001 RHODETTE R. EVANGELISTA", "1506-001 REYMOND M. BAYANI", "1506-002 RONALDO D. SY", "1510-001 EDEN L. CACA", "1510-002 RACHEL ANN M. CAJIPE", "1511-001 EDWIN S. MENDOZA", "1511-002 OLIVER A. MERIN", "1602-002 JONAMAY D. BENEDICTO", "1605-001 IMEE M. ARANAS", "1605-002 ANTHONY B. ATIENZA", "1606-001 GAYLORD A. BOHOL", "1607-001 IRISH D. LAJARA", "1701-001 AIDA D. ABESADA", "1701-003 CRISLYN V. AMANO", "1701-004 DESARIE V. ANDAL", "1701-005 SHELA-MAE M. APIL", "1701-007 ERMILYN B. OBNIALA", "1701-008 JEFFER B. BASAÑEZ", "1701-009 MYLENE M. BAULITA", "1701-014 DONNA O. CARAIG", "1701-016 JO-ANN MAY J. CARRILLO", "1701-020 MARY ROSE R. COMEL", "1701-022 JESTINE G. DE GUZMAN", "1701-024 ALELI APPLE R. DIAZ", "1701-025 ANIELYN D. VALENZUELA", "1701-065 CHRIS S. DIMAILIG", "1701-027 NESHILL A. DUMAGUING", "1701-028 ME-ANN A. ESPENILLA", "1701-029 SHIELA F. MALABUYOC", "1701-030 ANABELLE D. FRANCA", "1701-032 KAZEL M. GRIETA", "1701-033 NERMA V. GUZMAN", "1701-034 WENDY ANN M. HABLA", "1701-035 KRESIA ANNE C. HERNANDEZ", "1701-037 JULIE ANN A. JARITO", "1701-040 MA. NIÑA D. LIPAOPAO", "1701-041 MERESA M. LIWANAG", "1701-042 GENEDEL Q. MACINDO", "1701-043 LYKA S. MAGALLANES", "1701-044 MILDRED M. CASTILLO", "1701-045 YOLANDA O. MALACAMAN", "1701-047 MARY ROSE A. MARQUEZ", "1701-048 LANIE M. MALVEDA", "1701-052 MA. THERESA C. PEÑA", "1701-054 ALYSSA JANE S. PORCALLA", "1701-057 CHERRIE V. RECTO", "1701-058 JESSA A. REYES", "1701-059 INAH CHENTI D. RICOHERMOSO", "1701-062 LESLIE T. SALONGA", "1701-063 GINALYN R. SALOZA", "1701-066 EMMANUEL  JORDAN SANDOVAL", "1701-068 MYLENE D. TENORIO", "1701-069 KAYCEE S. UMALI", "1701-070 WILMA A. VALDEZ", "1701-071 DIANNE CAMILLE I. VILLANUEVA", "1701-075 MARY JANE B. MACAUBA", "1702-000 SATORU TOYOSHIMA", "1704-001 ELSA D. ACEDERA", "1704-006 LANNIE C. DELOS REYES", "1704-007 LILYBETH P. FAMODULAN", "1704-009 JAYVEE CLAIRE P. LOTERTE", "1704-011 JEMILYNE D. MATERUM", "1704-012 GENEROSE MAY P. OPEÑA", "1704-015 AIZA J. PROHIBIDO", "1704-017 SHIRLEY V. SAGUIT", "1704-018 REXCIE F. SANGALANG", "1704-019 RITCHELLE B. TAGLI", "1704-020 MARIAN A. LUMBRES", "1704-021 ARIANNE MEXIS FLOR P. ARRABE", "1704-022 MARIA KATHERINE B. BALAZON", "1704-023 JILLIAN B. ARANIEGO", "1704-024 GERLYN N. BIHIS", "1704-028 BUENA M. SERRANO", "1704-029 JOYLYN P. PETATE", "1704-030 MELISSA M. REYNA", "1704-032 CHARLENE L. VALENCIA", "1705-001 ANGELIQUE Z. ANTONA", "1705-002 RHINA ROSE V. CANTA", "1705-003 MAE JERAMY M. DE LA TORRE", "1705-006 DAISEREE E. VIERNES", "1705-007 KIMBERLY ANN G. BAROTILLA", "1705-008 RECHELLE B. GINES", "1705-009 HAZEL ANN G. ALVARIO", "1705-010 MARICAR J. PANGANIBAN", "1705-012 EDENROSE B. MAGISTRADO", "1705-013 ANGELIE O. MALIGALIG", "1705-016 WHYRINE C. MARANAN", "1705-019 LOTA M. LARA", "1705-021 VIENALEN B. ORPILLA", "1705-022 MARIA KAYE L. PEREZ", "1705-023 JHUNALYN R. PAGMANOJA", "1705-025 ALVIN J. ARAÑES", "1706-001 ESTEFHANNY E. LIM", "1706-002 MELODY S. CRUZAT", "1707-003 LOIDA B. LIBIRAN", "1707-004 ANGELYN I. ROXAS", "1708-001 CHEENEE M. DIAZ", "1708-002 MA. AILEEN F. AGUILA", "1708-003 FATIMA M. ATIENZA", "1708-004 ERICK M. BUNDALIAN", "1708-005 ARIEL C. GUEVARRA", "1709-001 MA. WELLA Q. CARBILLON", "1709-002 MYLEN A. IRINGAN", "1709-003 JOVIELYN P. MAGPANTAY", "1709-004 ARLENE S. ROXAS", "1709-006 CLARK ANTHONY B. VILLANUEVA", "1709-007 NICA C. VILLARIEZ", "1710-002 APRIL F. MORCOZO", "1710-003 MARIFE C. SAN JOSE", "1710-004 REZIEL C. ASUNCION", "1710-005 JESSAMIN C. CANDIDO", "1710-007 RICA D. CLEMENSO", "1710-010 CHARLENE M. EVANGELISTA", "1710-011 EDMAR D. FERRER", "1710-012 MARY ANN O. GUEVARRA", "1710-013 PAMELA M. HERNANDEZ", "1710-014 APRILYN T. JARIÑO", "1710-018 ZERLIE ANN T. PALAAC", "1710-020 JINKY T. SOLLER", "1710-021 CRISHELDA B. SUGANOB", "1711-002 JENNIFER S. BUNO", "1711-003 GERALDINE V. BUNSOL", "1711-005 MAILA M. EVANGELISTA", "1711-007 AIRENE B. FORBES", "1711-008 EDELIZA G. JAEN", "1711-010 RACHELLE H. RODRIGUEZ", "1711-011 JEAN M. ROMAN", "1801-002 VENGIE P. AGRIPA", "1801-003 HAZEL M. ALMAREZ", "1801-004 GLADYS L. ALMENDRAL", "1801-005 LIZA M. ALVAREZ", "1801-008 ARJELEN C. BARNEDO", "1801-009 CLARINE C. CALIVO", "1801-010 BABY JOY R. CARLOM", "1801-012 MARY JANE M. DESTA", "1801-013 GELLI E. FLORESCA", "1801-018 KAIRA MAE V. GARCIA", "1801-019 DANICA C. LASAT", "1801-021 NEMESIA H. MALABANAN", "1801-024 ROSELYN I. MANAY", "1801-026 ABEGAIL M. MEERA", "1801-028 DORRIS A. OBRADOR", "1801-029 RACHELL P. PILAPIL", "1801-030 IVY MAY L. PIÑON", "1801-031 ROSE-BELL B. SANGALANG", "1801-032 JASTINE A. TAPARANO", "1801-033 JIRLEEN M. TINA", "1802-001 CATHERINE L. DELA PEÑA", "1802-003 ALIZA V. BUENDIA", "1802-004 NELIZA G. CASTAÑARES", "1802-005 MICHELLE P. DELA CRUZ", "1802-006 ROSE B. DIMAYUGA", "1802-007 DANICA P. ILAG", "1802-008 MARIFE C. ILAO", "1802-009 ANGEL MAE U. INLOCNA", "1802-011 REA MAY M. LEYNES", "1802-014 MICHELL A. LOSA", "1802-015 RODELISA G. MARANAN", "1802-016 JULIE P. MILLA", "1802-019 MARY ANN S. ONAN", "1802-021 PRINCESS HAZEL G. VASQUEZ", "1803-001 AILA MAY A. FELESMINO", "1803-003 MA. CHRISTINE T. ALEGRIA", "1803-004 GERALDINE F. AQUINO", "1803-006 HAZEL CARALIPIO", "1803-007 JINKY D. CONSEBIDO", "1803-008 JOY O. FABELLA", "1803-009 RONEILYN G. GOLPEO", "1803-010 KAREN F. NORTE", "1803-011 RIHCA MEA C. SASOTA", "1803-016 ROWENA B. BRIONES", "1803-017 JULIET M. CASTILLO", "1803-019 LORENA J. JASARENO", "1803-020 JASPER C. JAVIER", "1803-021 RHAE MAE R. NAGUIMBING", "1803-022 RINA L. SOQUIAT", "1803-023 LIENCIE B. VELASCO", "1803-000 AKIO SHIMAMURA", "1805-002 CHARIE N. MOSCA", "1805-003 IRENE KRISTY S. TOREJAS", "1805-004 AIZA D. ROSALES", "1805-006 VINCENT ELOI A. MAGBOJOS", "1805-007 FLOREFE R. MALVEDA", "1805-008 MARICAR Q. MEERA", "1805-009 VANESSA E. QUIOZON", "1805-010 RACHELLE C. URGELLES", "1806-001 MERY R. DEDUYO", "1806-002 LANIE G. ESTAÑOL", "1806-003 JUDITH C. ORTEZ", "1806-005 SHERYL A. SICLON", "1806-006 LOVELYN A. TAYCO", "1807-002 HARRY J. TAÑEGA", "1807-003 LAARNIE L. ABANILLA", "1807-005 MARGIE A. BACULPO", "1807-006 MARY ANN L. BADASAN", "1807-007 JONALYN S. BALBA", "1807-008 ROSE MARIE A. DIENTRE", "1807-011 ELSA C. ESCALONA", "1807-012 GLAEZLE E. MANALO", "1807-015 JOAN L. ANDAYA", "1807-016 HAZEL E. MACASADIA", "1807-018 LEA R. MASUNCA", "1807-019 CLARIBEL H. MENDOZA", "1807-021 ERICA O. HABLA", "1807-022 ROCELYN B. OREA", "1807-023 VIA I. SANTOS", "1807-024 MA. DONA S. SEDUTAN", "1807-025 SHERLYN G. TUMAMBING", "1512-000 KENYA FUNAKI", "1809-001 NERISSA L. ANDES", "1809-002 MELANIE P. ARIOLA", "1809-003 DIVINA M. BALUDDA", "1809-004 MARJORIE T. BOLANTE", "1809-005 ARMANDO C. BUERE I", "1809-007 DIANE H. CASTILLO", "1809-008 NERISSA D. CATAPANG", "1809-009 ANGELICA B. CAYETANO", "1809-010 MARRY GRACE P. ESTEVES", "1809-012 MARY ANN S. IGNACIO", "1809-013 CRISTINA R. LEPALAM", "1809-015 DIANA JOY S. MADAMBA", "1809-016 JONATHAN R. MAGLUYAN", "1809-018 JACKIE LOU G. MALACAPO", "1809-020 GINALYN C. MANALO", "1809-022 ROWENA S. ORTEA", "1809-023 PERLA H. PACONIO", "1809-024 JASMINE D. PASOQUIN", "1809-026 CHERILYN M. ROBEL", "1809-027 INTISHAR M. SAGOSARA", "1809-028 MARY GRACE T. RIEGO", "1809-029 DANIEL BALVINO D. DALISAY", "1809-031 RICHARD M. PEÑALBA", "1809-030 VIRULYN P. GEMOTRA", "1809-032 PRECIL P. PRADO", "1810-001 REINA GLADYS M. AGUILA", "1810-003 SHELLA C. RECIO", "1810-007 SHENNA A. PADILLA", "1810-008 ERICSON S. VERGARA", "1811-002 JAY-AR H. AREGLADO", "1811-004 DINA T. BANDOJO", "1811-005 MARJORIE C. BAUTISTA", "1811-008 BRYAN ANTHONY B. BURCE", "1811-010 KAREN G. CAMPANA", "1811-011 GINALYN D. CANLOBO", "1811-014 ANGELICA G. COMIA", "1811-015 LENIE M. DELA PEÑA", "1811-018 GINA A. GARCIA", "1811-019 PRESCILA E. GONZALES", "1811-020 ZAIDA C. LOPEZ", "1811-021 ALEXA DANIELLE C. MANALO", "1811-022 SARAH D. MARQUEZ", "1811-023 JUVY B. PACRES", "1811-024 BRENEELYN O. PASOOT", "1811-025 MA. ELOISA D. RESURRECCION", "1811-026 CHENELYN L. SAJO", "1811-027 MARIA NIEVA G. SOFRANES", "1811-028 MARICAR M. TABAT", "1811-030 EDWIN S. SANCHEZ", "1811-031 JHON B. LONTOC", "1901-001 NOEMI D. DE VILLA", "1901-002 GLEN TRISTAN T. RICAMARA", "1901-005 LOVELY JOY D. ANAYA", "1901-006 RUTH P. ANCIADO", "1901-007 MARY GRACE L. BALURAN", "1901-009 SHARAJANE C. BRIZ", "1901-011 DYAN D. CANIETE", "1901-012 MARY JANE C. CUALQUERA", "1901-014 JAYBETH A. DIMAPALITAN", "1901-015 RIZELLE E. SACRISTIA", "1901-018 KAREN S. GONDRANEOS", "1901-019 MELBA M. MEJILLA", "1901-020 REBELYN C. MENDOZA", "1901-021 MARINETTE B. MONZON", "1901-025 MICHELLE ANN L. RAMOS", "1901-026 CARREN V. RAPSING", "1901-027 EFRELYN B. REPIA", "1901-028 RICHARD D. RIVERA", "1901-029 ELLA MAY I. SAGARIO", "1901-031 JUDY ANN C. UBALDO", "1901-032 LESLIE E. VISTA", "1901-033 NORIEL B. AQUINO", "1902-002 RUTCHEL O. ANDA", "1902-005 JENNIFER B. BELMONTE", "1902-006 MARY ANN BINLAYO", "1902-008 MARISTELA S. DE CASTRO", "1902-011 GEANETTE V. EBARDONE", "1902-012 MARIALYN ELARDO", "1902-014 ROSALIE M. MANGABAN", "1902-015 MANILYN A. MEDILLO", "1902-016 JOANNE L. OLA", "1902-017 BRENDA T. RIVERA", "1902-018 JOYFOLL E. RIVERA", "1902-019 JONALYN M. SILANG", "1902-020 REYNALDA V. TERO", "1902-021 JESSICA MARIE A. VILLANUEVA", "1902-023 ROSLYN H. VILLARBA", "1903-001 JOVELYN G. RENDORA", "1903-002 LAWRENCE B. ORINDAIN", "1903-003 ARIANE G. ARANTE", "1903-004 EMALYN D. ARMENTA", "1903-005 RONALYN P. AVETRIA", "1903-007 LYRA N. HULAR", "1903-008 GENNALYN T. ODOÑO", "1903-009 RUBY B. RELANES", "1903-011 CHRISTIAN M. NEDULA", "1903-012 MA. CRISTINA D. PEREZ", "1904-001 MARY GRACE M. ASENCE", "1904-002 MARY ANN B. LLAMERA", "1904-004 ANGELISA B. DE VERA", "1904-006 VISITACION E. GUMIRAN", "1904-007 MARY ROSE N. MAGSINO", "1904-010 JOAN R. MUSA", "1904-012 SHERLYN D. PAGUNSAN", "1904-015 CRESTIE G. SANTONIA", "1904-016 GEL D. SARNE", "1904-018 FRANCE JOYCE S. MANALO", "1905-001 MARRY ANN D. BATHAN", "1905-004 LORIE ANN U. CONSTANTINO", "1905-005 JOANNA E. DE LUNA", "1905-006 GEMMALYN O. DE TORRES", "1905-007 JOANNE V. DE LOS SANTOS", "1905-008 EMILY G. FLORES", "1905-010 PRECY M. MEDINA", "1905-011 APPLE M. MONTEJO", "1905-012 ANNA MARIE C. MORALLOS", "1905-015 JEYSA A. TARGA", "1905-016 MA. THEA R. TUGNAO", "1905-017 MA. CHRISELDA D. VILLEGAS", "1905-018 CARLA O. LUNA", "1906-002 CHARIE DELL S. CUADRO", "1906-003 ANABEL B. GALIT", "1906-004 JANELLE M. INDITA", "1906-005 RITCHEL ANN M. LUMANG", "1906-006 JUDY ANN MANGUSTA", "1906-009 ELLIE ROSE C. PALANG", "1906-010 RICHELLE M. REDOMA", "1906-011 DIANA LYKA R. RODIL", "1906-012 SANIKA MAE D. ROSAS", "1906-013 ROMALYN D. SOLOMON", "1907-002 DONNABEL V. CORNERO", "1907-003 MARIA CHERAMIE B. DOMINGO", "1907-004 REINA JOHAN R. GUTIERREZ", "1907-005 ANGELICA D. MADRIDANO", "1907-006 ROVELYN A. MARISTELA", "1907-007 ELSA A. PAGLINAWAN", "1907-008 KHEYNETH P. RODRIGO", "1907-009 MERILYN E. ROMAY", "1907-010 JESSICA V. SALVA", "1907-011 RECELYN E. SAN BUENAVENTURA", "1907-012 CHRISTINE MAY C. SERNA", "1907-013 NESTOR JR. T. TABIGNE", "1907-014 JANETH A. BELETARIO", "1907-015 NOVA S. CABALLERO", "1908-001 EVA L. ANTONES", "1908-002 ROSALIE C. ASI", "1908-003 JENNY ROSE M. ATIENZA", "1908-004 JUDITH N. AVILA", "1908-005 CHRISIEL M. BANTOC", "1908-006 JUDY ANN BAYLOSIS", "1908-007 JOYCE ANN D. CAMBAL", "1908-008 APOLONIA G. CUETO", "1908-009 BERNALYN G. DEQUITO", "1908-010 KIM LYNARD Y. HERNANDEZ", "1908-011 ROSE ANN C. HOLGADO", "1908-012 JESSICA T. LOZADA", "1908-013 REGINE JOY R. MUNDA", "1908-015 KATHLEEN G. PORRAS", "1908-017 CATHERINE A. ROBERTO", "1908-018 ANALYN D. SAUT", "1908-020 MARY GRACE E. SIONILLO", "1908-022 MAYBELYN P. TENORIO", "1908-023 RUTH O. YBARDOLAZA", "1908-025 MARY JANE M. PERALTA", "1908-024 MARICEL M. CATAPANG", "1909-001 MARENELL G. ALAYAN", "1909-002 BENILDA V. ARCIAGA", "1909-003 SHELLA M. BREDES", "1909-004 IRENE P. DELOS SANTOS", "1909-005 EDEZA M. DIZON", "1909-006 RODEL R. DOCTORA", "1909-007 MAITA S. LORILLA", "1909-008 AILYN N. MAGTIBAY", "1909-009 CHRIS ANN L. MAQUIÑANA", "1909-010 ANDREW C. NACARIO", "1909-011 JAN CAMILLE M. OJO", "1909-012 ARLENE C. ORLINA", "1909-014 JENNA MARIE S. PESIT", "1909-015 JUDYLYN M. ROMANO", "1909-016 RASHERL G. SENO", "1909-017 MICHAELA ROSE VIRTUCIO", "1909-018 JOBYLYN ZABALA", "1909-019 DAISY ANN D. DADULLA", "1910-003 MATETH N. ALISOSO", "1910-004 MELISA R. AMPER", "1910-005 MARICAR H. ARGAÑOSA", "1910-007 MELODY M. BAUTISTA", "1910-008 BERNADETTE B. BRIBON", "1910-010 CRISELDA T. CABANES", "1910-011 GERALDINE L. CAFE", "1910-012 APRIL G. CAMARA", "1910-013 MADELYN J. COLLADO", "1910-014 DYRINE A. DE CASTRO", "1910-018 MARK ANTHONY R. LIGAS", "1910-020 JENNY ROSE M. MANALO", "1910-021 YANIE N. NOYNAY", "1910-022 JOMIELYN I. TAN", "1910-023 JOSIE M. VERA", "1911-001 CHARITY C. BORJA", "1911-002 DECEREAN L. GARCIA", "1911-003 ANABEL P. LAURETA", "1911-004 SARAH JANE D. MAILOM", "1911-005 SALLIE M. MENDOZA", "1911-006 ROSELYN P. MIOLE", "1911-007 DIANNE M. OPEÑA", "1912-001 MARK RYAN R. CUYAPEN", "2001-007 IMELDA G. MARTINEZ", "2001-018 JOVI CLAIRE N. SANGALANG", "2001-005 RAMILYN N. DELA VEGA", "2001-016 JENNY G. BUTIAL", "2001-002 JESSA C. ATIENZA", "2001-006 JOLINA T. LANOZA", "2001-013 MIALYN C. TANDOY", "2001-010 LEIZA D. PACURIB", "2001-008 MARIEL D. MENDOZA", "2001-015 NELYN S. BALONDO", "2001-011 DAHLIA N. REDULA", "2001-009 MERY ROSE P. MONTERO", "2001-014 MICHA-ELLA A. TRIPULCA", "2001-017 CLAIRE B. DE LOS SANTOS", "2001-003 RALYN R. AGOJO", "2001-004 HAZZEL M. BAUTISTA", "2002-001 LESLIE E. MALABANAN", "2002-003 DONNA M. AGUILAR", "2002-004 OTCHE I. BALDOZ", "2002-005 JAIRO D. CABRAL", "2002-006 JESSA MAE R. CAMPO", "2002-007 GE-ANNE L. CANDELARIA", "2002-008 MANELYN D. DATINGUINOO", "2002-009 FRANCE JOYCE G. DIMACULANGAN", "2002-010 MARIBEL B. ESPLAGO", "2002-011 JANICE P. GARAY", "2002-012 JESSA D. HUBILLA", "2002-014 AMY L. MACASAYA", "2002-015 MARIA MERCEDES R. MAPANAO", "2002-016 AILEEN S. MONTECILLO", "2002-018 PRINCESS ODETTE NOGALES", "2002-019 KATHERINE S. OÑATE", "2002-020 AMOR H. PAHILANGCO", "2002-022 HAZEL ANN D. PIÑON", "2003-001 JOSEPH B. PUEBLO", "2004-002 GEMMA U. COROT", "2004-001 MARIAN A. BAGASINA", "2004-003 ANA MAE D. DELA GENTE", "2004-004 CHRISTIAN MAE H. DONAYRE", "2004-005 GRACE C. ESPIÑA", "2004-006 GIL R. LAPARAN", "2004-008 JESABEL P. MARTINEZ", "2004-009 MORIEL B. MENDOZA", "2004-011 JONEL P. OBAÑA", "2004-012 MAY MARIE B. OGORIDA", "2004-013 ARLENE P. PELAEZ", "2004-014 JESSICA M. PUNTANAR", "2004-016 CHELLO C. VELEZ", "2004-017 PRINCESS ROSE S. ZULUETA", "2005-003 REALEE M. ABANILLA", "2005-004 MARY GRACE L. AGUILAR", "2005-005 CHARMANE ROSE ANN D. ALAP", "2005-006 MERCY C. ALVAREZ", "2005-007 KRISTEL P. ATIENZA", "2005-008 SHIELA MAE E. AUSTRIA", "2005-009 MELANIE AZUELO", "2005-010 MARLYN M. BADILLO", "2005-011 GERALDENE V. BADO", "2005-012 EDELYN B. BALDE", "2005-013 LEZLIE C. BASCONCILLO", "2005-014 GERALDINE A. BIÑAS", "2005-015 CLAIRE R. CALIMOTAN", "2005-016 ANGELICA C. CASAPAO", "2005-017 LENY M. CASAPAO", "2005-018 SHERLYN P. CAYMO", "2005-019 MEAN G. DE LEON", "2005-020 LORENA L. DE LEON", "2005-021 RIZA L. DEL MUNDO", "2005-022 NIKKI M. DELOS SANTOS", "2005-023 CARMAY B. DILAO", "2005-024 CRISTINE M. DOMONDON", "2005-025 GLADYS L. GAMBOA", "2005-026 REGEN J. HOMERES", "2005-028 LALAINE J. MANCAY", "2005-029 STEFANIE MONTEVERDE", "2005-030 NIÑO TRISTAN H. NAMIA", "2005-032 EMERSON L. OLANDA", "2005-033 MA. ROSENDA G. ORTIZ", "2005-034 JOBELLE T. ORUGA", "2005-035 RHEA D. ORUGA", "2005-036 JAMIE G. PARSA", "2005-037 JONALYN T. PLATEROS", "2005-038 LERICKA JANE M. REDONDO", "2005-039 CLARISSA E. ROLL", "2005-040 MIA ROSE R. SALES", "2005-041 MA. TERESA M. SAN RAMON", "2005-043 JESSICA G. VALENZUELA", "2005-044 HEIDY O. VILLANUEVA", "2005-045 WENDY P. VILLANUEVA", "2005-031 GENEVA Z. NIMEDES", "2005-001 CHAD S. BELEN", "2005-002 ELDRIN BRYLLE M. LATAYAN", "2005-046 ANNEBELLE B. AYOP", "2005-047 ROBERT ANTHONY T. ALCANTARA", "2005-048 GLEZEL MAE E. GILO", "2005-049 ANGELICA M. NIEVA", "2006-001 JOYMEE Y. ABOT", "2006-002 FHERLYN E. AGUNOD", "2006-004 EVA MEY O. CASIL", "2006-005 MA. ALEXANDRA P. CASTILLO", "2006-006 BIANCA CAMILLE R. CATALUÑA", "2006-007 ROSEMARIE L. ESPERAGOZA", "2006-008 MARINEL G. GALAMGAM", "2006-009 JOVEL A. MALVEDA", "2006-010 ROSALYN L. MANJARES", "2006-011 MERRY CRIS B. MENDOZA", "2006-012 MILDRED E. MONTECILLO", "2006-013 MELODY C. REPOMANTA", "2006-015 COLENNE D. SUAREZ", "2006-016 SHEINA JOY D. TABLAR", "2006-017 JENNELYN N. TALARO", "2006-019 JAYSON M. TOLENTINO", "2006-020 CHRISTIAN M. TORRES", "2006-021 MEDELYN B. DELOS REYES", "2007-002 IVY M. CUETO", "2007-003 REYCHELLE B. DAVID", "2007-005 CINDY V. INCOMIO", "2007-006 MANILYN C. INOFRE", "2007-007 JESSALYN M. MAGPANTAY", "2007-008 JONALYN B. MOTILLA", "2008-001 LORNA B. BALOBAR", "2008-002 DAISY L. CAILAO", "2008-003 AIRALYN L. DIOLAS", "2008-005 DOROTHY D. PAGO", "2008-007 ARIEL Q. VILLAFLOR", "2009-002 EFRAIM JOSHUA B. MALIBONG", "2009-015 LIZA C. PASTRANA", "2009-003 MA. KEY KRISTINE E. ABELILLA", "2009-004 HANNIE MAE N. ALIVEN", "2009-005 MARY JOY M. CADAY", "2009-006 MARY GRACE A. DESEPEDA", "2009-007 MARY ANN C. DORADO", "2009-008 SHYRINE-ANN A. EREVILLA", "2009-010 JENNILYN H. HALUM", "2009-012 CHERRY H. LITERAL", "2009-013 JINKY C. MALIGO", "2009-014 MARIEFLOR M. MIRABALLES", "2010-003 RHONALYN P. CARANDANG", "2010-004 HYZEL M. CASTILLO", "2010-005 ESPERANZA A. CERVANTES", "2010-006 RUBY-ROSE ESTOR", "2010-007 MAYLEN L. ESTRADA", "2010-008 JOANA LESLIE M. ESTRELLA", "2010-009 ANALYN JEPOLLO", "2010-010 CORAZON Q. LEMUS", "2010-011 SHIENY C. MACARAIG", "2010-012 GLADYS G. MARCELO", "2010-013 MELDRID S. MASILANG", "2010-014 ZYLINE T. MENDOZA", "2010-015 JESSA A. MILITANTE", "2010-016 LEA T. PAMPLONA", "2010-017 ED CHRISTINE T. RUNAS", "2010-018 ALMA M. SABUITO", "2010-019 JULIET L. VILLANUEVA", "2011-001 JEVELYN R. ARGAÑOSA", "1907-000 YUKI YAMAMOTO", "2011-004 DOREEN MAY T. ANINGAT", "2011-005 MICHELLE S. ARCILLA", "2011-006 MAURICE E. ASIA", "2011-007 MA. ERICA B. BARJA", "2011-008 APRIL P. BELOLO", "2011-009 RYA PEARL S. CAGAT", "2011-010 MARY GRACE L. DEL PILAR", "2011-011 JOCELYN V. DELA CRUZ", "2011-012 CLEALYN A. FABRERO", "2011-013 IVY T. GERNALIN", "2011-014 MARY CRIS B. LITANA", "2011-015 FLORDELYN M. LUSTANIA", "2011-016 MAYRIN M. MACASADIA", "2011-017 LESLIE V. MAGBUJOS", "2011-018 DHENNY M. MANIPOL", "2011-019 JOAN L. MANONGSONG", "2011-020 JOCELYN A. MARQUEZ", "2011-021 REIGINE A. MENDOZA", "2011-023 GEMMA P. PALO", "2011-024 MADEL Y. PAMPLONA", "2011-026 ANGEL D. PAR", "2011-027 MARY JOY V. SEROMA", "2011-028 CAREN H. TORALBA", "2011-030 ROSE LYN M. VILLEGAS", "2101-001 AUGUSTINA F. ALERTA", "2101-002 JOAN B. BARRION", "2101-003 RENIEL J. BOTE", "2101-004 KIRK BRYAN B. CABAREÑO", "2101-006 EDMON C. DE GUZMAN", "2101-007 EDUARD JR. C. GUZMAN", "2101-008 JULIET R. LUGO", "2101-009 GERICO C. LUZ", "2101-010 RANDY E. NOGALES", "2101-011 ROSIE E. NOGALES", "2101-012 MICHELLE M. OSINSAO", "2101-013 EVERLYN H. CADANG", "2101-014 CECILE CENTENO", "2101-016 CAMILLE JOY D. DIMAANO", "2101-018 LORA F. LAYDEROS", "2101-019 ZENAIDA A. LLORCA", "2101-020 EVELYN S. LUMANGLAS", "2101-021 BABY SANILYN B. MALIPOL", "2101-022 GENELYN U. NITURAL", "2101-023 MARY ROSE A. NUÑEZ", "2101-024 JAYZELLE C. OLIVERA", "2101-025 CHERRY ANN A. RAMIREZ", "2101-026 MA. SHAIRA H. REY", "2101-027 CRISANTA A. RODELAS", "2101-028 ADONISA D. TAYCO", "2102-000 NORIKO NISHIDA", "2102-001 IRENE T. AALA", "2102-002 GLADYS A. ABEJUELA", "2102-003 BERNADETTE P. ADORNADO", "2102-004 ANALYN G. AMPUNIN", "2102-005 MARIA RICA H. ANDINO", "2102-006 ANGELICA E. ARIOLA", "2102-007 ALLYZA C. ARSENIAS", "2102-008 MARICAR A. ASUNCION", "2102-009 MELANIE R. AYCOCHO", "2102-010 MELBEVS V. BALEJONDA", "2102-011 JHONNALYN D. BAUTISTA", "2102-012 MARA M. BUSTAMANTE", "2102-013 RESEL MAE E. CARALIPIO", "2102-014 MARITES C. CASTILLO", "2102-015 MARY GRACE M. CLOSA", "2102-016 MARY GRACE B. CUDANIN", "2102-017 MARY ANN D. DELLOVA", "2102-018 MORRIEL ANN C. DIMEN", "2102-019 RUBIE V. DIOMAMPO", "2102-020 SAIDY L. DIZON", "2102-021 MARIA VANESSA R. DONGON", "2102-022 MA. GEMMA SARAH C. ENRIQUEZ", "2102-023 JUDY ANN EVALDEROZA", "2102-024 NECCA ANDREA R. FERNANDEZ", "2102-025 SHEILA S. GARCIA", "2102-026 MARJORIE B. GUERRERO", "2102-027 MYRA N. JAMISOLA", "2102-028 MARIBEL E. KALIBUSAN", "2102-029 LAYLINE A. LACOPIA", "2102-030 JEANCEL O. LANOZA", "2102-031 PRINCESS ANN A. LLAMES", "2102-032 JENALYN D. MANABAT", "2102-033 KAREN M. MARTINEZ", "2102-034 ROCHELLE MAUHAY", "2102-036 VERNADETH N. MONTALES", "2102-037 JOVELYN M. PEREZ", "2102-038 LORENA M. QUILESTE", "2102-039 JOONAMER T. ROBEJES", "2102-040 JUDY ANN G. SANCHA", "2102-041 ABEGAIL E. SANCHEZ", "2102-042 ROCHELL SANCHEZ", "2102-043 MICHELLE G. TANUYAN", "2102-044 ROXANNE M. VALENZUELA", "2102-046 ANDRIAN J. ZARA", "2103-000 HIROMI N. GAVILENO", "2103-001 JOSEPH M. REAMBONANZA", "2103-002 ODETTE P. ALCANTARA", "2103-003 MARIANNE A. ANDALLON", "2103-004 LORIEGENE B. CASTILLO", "2103-005 CHENEY MAY A. DELA PENA", "2103-007 VEVERLY V. ESTORES", "2103-008 MARY JANE T. HERNANDEZ", "2103-009 APPLE L. NACION", "2103-010 JENIE F. OBRADOR", "2103-011 IRICA D. PADILLA", "2103-012 RENILYN C. PASCUAL", "2103-013 LORRAINE G. PILI", "2103-014 ELLEN MAE P. PORRE", "2103-015 BABY JANE P. SABIGAN", "2103-016 LAARNI G. SAGUN", "2103-017 JOVELYN L. SALVACION", "2103-018 JESSICA B. SILABA", "2103-019 JOVY G. SOLAYBAR", "2103-020 CARMINA L. VILLANUEVA", "2104-003 MARIEANNE G. BARRION", "2104-012 NIÑA F. ABANID", "2104-013 NORIE B. ANDALIS", "2104-014 MARICON P. CASARINO", "2104-015 ROSE ANN L. CATINA", "2104-016 MA. CARINA CHARMAINE PIA", "2104-017 JANICE M. CRUZADO", "2104-018 CARREN C. DAÑO", "2104-019 MARY ACEL DERELO", "2104-021 KIM ABIGAIL V. GARCIA", "2104-022 MARY ANN P. GRAJO", "2104-023 ELOISA M. GREÑO", "2104-024 JAEMORE HERNANDEZ", "2104-025 SHEILA MARIE R. LORENZO", "2104-026 KHIMVERLY M. LUCERO", "2104-027 ARTH DOMINIQUE T. LUCILLO", "2104-028 MARY JANE S. MAASIN", "2104-029 JAYSON T. MARTINEZ", "2104-030 LESLIE ANN O. MERCADO", "2104-031 SHEENA MARIE C. MESIAS", "2104-032 KAYE ANN MAY L. MESOLANIA", "2104-033 EVELINDA L. OBSINADA", "2104-034 ZYRA MAY G. OJAS", "2104-035 LOREZA O. PACINOS", "2104-036 JERAH G. PANALIGAN", "2104-037 MARIMAR G. REJOSO", "2104-038 CATHERINE D. RID", "2104-039 CARLA M. ROSAS", "2104-040 EMERIE SABASCAR", "2104-041 MARY ROSE A. LINATOC", "2105-001 ROSHELL L. ARELLANO", "2105-002 BERNADETH C. BECHAYDA", "2105-003 JOY C. CALAGUAN", "2105-004 JERLYN C. CARINGAL", "2105-005 JAYNALYN B. SUNE", "2105-006 JENNELYN A. ZATARAIN", "2106-001 JOVELLY A. APUYAN", "2106-002 ANANEL A. BAYO", "2106-003 APRIL JANE T. BICO", "2106-004 GEMMALYN P. CARIÑO", "2106-005 JILL E. DIMAPILIS", "2106-006 MARVIN Q. MACATANGGA", "2106-007 JINGKIE L. MAGSINO", "2106-008 CAMILLE E. MALIPOL", "2106-009 ANGELICA C. MANIMTIM", "2106-010 HAYDEE A. MENDOZA", "2106-011 RANDYLEN P. MONDIGO", "2106-012 MIRASOL B. MORONG", "2106-013 ROSEMARIE U. SAN LORENZO", "2106-014 ANGELICA A. SARMIENTO", "2106-015 LALAINE P. SARMIENTO", "2106-016 AIZA W. TOLENTINO", "2106-017 MARK AERON F. CARANDANG", "2106-018 LEA MAE N. MANGUBAT", "2106-019 KYAN KHEN L. SY", "2106-020 MARY KARLIN JOYCE C. TANO", "2106-021 MACKY B. VILLANUEVA", "2106-023 KRIZHA SOL C. ANO", "2106-022 JAN VINCENT M. PEREZ", "2106-024 KRISTINE M. GARCIA", "2107-001 CATHERINE A. SIM", "2107-002 MELODY A. CONZON", "2107-003 ALELYN C. LUNA", "2107-004 MARY ANN A. MAOMAY", "2107-005 ROVELYN A. MENDOZA", "2107-006 REYMARIE L. OLIVERA", "2107-007 ANGEL E. ALINGATO ", "2107-008 REYNALYN C. BINGHOY", "2107-009 MADELINE M. GARCIA", "2107-010 KIMBERLY U. NATANAUAN", "2108-001 MA. LESLEY L. DE LUNA", "2108-002 SHYRA B. ESTOYA", "2108-003 ANTONINO T. GAVILENO", "21-001 SHANA ARGANDA", "21-002 ARA MAE R. BARRANDA", "21-003 RICA MAE C. MALAIBA", "21-004 JOHN KENNETH NAVARRO", "21-005 MARK JOSHUA P. PAZ", "21-006 KYLA MAE  T. RUBIO", "21-007 JOHN KENNETH O. RUBION", "21-008 MONICA V. ZATA", "2108-004 AIVY JOY A. CANTA", "2108-005 AILEEN L. CENTENO", "2108-006 MARRY JANE C. DE CHAVEZ", "2108-007 RIZZA G. ESMINO", "2108-008 BEVERLY F. FRONDA", "2108-009 LENDY M. GARGOLES", "2108-010 PAMELA A. LANGA", "2108-011 NECY T. LONTOC", "2108-012 JENNELYN C. MAGADIA", "2108-013 MYLENE D. MAGNO", "2108-014 NERISSA A. MALABANAN", "2108-015 MORENA C. MANALO", "2108-016 JOLINA D. PADUA", "2108-017 JIETH JOY D. PONCIANO", "2108-018 ANGELA P. REQUISO", "2108-019 JOSEPHINE IRENE ROJO", "2108-020 LIEZEL N. SABENIANO", "2108-021 JOY A. SONIEDO", "2108-022 CATHERENE V. TERCIAS", "2108-023 CRISTINE ELAINE A. VARELA", "2108-024 JOYME A. VILLANUEVA", "2108-025 MA. LOVELY P. ARIOLA", "2109-001 KATHERINE ANDAY", "2109-002 ALAYKA MAE A. ANGULO", "2109-003 JENNIFER S. CRUZ", "2109-004 KAYE ANN M. DORIA", "2109-005 DAISY H. FABRIQUEL", "2109-006 LADYLIN F. GABAY", "2109-007 JUDY ANN G. LAGUSTAN", "2109-008 MARY JOY R. MAYOR", "2109-009 DIVINA NOCOS", "2109-010 JANINE U. OLIVA", "2109-011 MARA JANE R. PARADERO", "2109-012 ALAISSA JOY G. PERMEJO", "2109-013 ROSALYN B. RAMOS", "2109-014 HIEZLE M. RIVERA", "2109-015 BEVERLY S. TURADO", "2109-016 MICHAEL D. UMALI", "2109-017 JEAN R. VARGAS", "2110-001 RENALYN T. RESABA", "2110-002 PRINCESS CAMILLE V. BAUTISTA", "2110-003 ERICA M. DOTE", "2110-004 EVALYN B. HAPIN", "2110-005 LARAH M. NAPAY", "2111-001 LIEZELLE P. MALDIA", "2111-002 MICHAEL R. BANABAN", "2111-003 JENNIFER D. BREIS"})
        Me.cmbEmploymentType.Location = New System.Drawing.Point(92, 31)
        Me.cmbEmploymentType.Name = "cmbEmploymentType"
        Me.cmbEmploymentType.Size = New System.Drawing.Size(491, 24)
        Me.cmbEmploymentType.TabIndex = 4
        '
        'lblLeaveType
        '
        Me.lblLeaveType.BackColor = System.Drawing.SystemColors.Control
        Me.lblLeaveType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLeaveType.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.lblLeaveType.ForeColor = System.Drawing.Color.Black
        Me.lblLeaveType.Location = New System.Drawing.Point(586, 5)
        Me.lblLeaveType.Name = "lblLeaveType"
        Me.lblLeaveType.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.lblLeaveType.Size = New System.Drawing.Size(88, 24)
        Me.lblLeaveType.TabIndex = 536
        Me.lblLeaveType.Text = "Leave Type"
        Me.lblLeaveType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbLeaveType
        '
        Me.cmbLeaveType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbLeaveType.Font = New System.Drawing.Font("Verdana", 9.5!)
        Me.cmbLeaveType.FormattingEnabled = True
        Me.cmbLeaveType.Items.AddRange(New Object() {"Sick Leave", "Vacation Leave", "Birthday Leave", "Marriage Leave", "Bereavement Leave", "Paternity Leave", "Offset Leave", "ECQ Leave", "Magna Carta for Women Leave", "Maternity Leave", "Half Day", "Undertime", "ECQ Leave (Quarantine)"})
        Me.cmbLeaveType.Location = New System.Drawing.Point(673, 5)
        Me.cmbLeaveType.Name = "cmbLeaveType"
        Me.cmbLeaveType.Size = New System.Drawing.Size(635, 24)
        Me.cmbLeaveType.TabIndex = 3
        '
        'frmScreenReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1314, 601)
        Me.Controls.Add(Me.rptViewer)
        Me.Controls.Add(Me.pnlTop)
        Me.Font = New System.Drawing.Font("Verdana", 8.5!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Name = "frmScreenReport"
        Me.ShowIcon = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Monitoring Report"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rptViewer As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents cmbDateType As ComboBox
    Friend WithEvents pnlTop As Panel
    Friend WithEvents lblStatus As Label
    Friend WithEvents cmbStatus As ComboBox
    Friend WithEvents btnClose As PinkieControls.ButtonXP
    Friend WithEvents btnClear As PinkieControls.ButtonXP
    Friend WithEvents btnGenerate As PinkieControls.ButtonXP
    Friend WithEvents lblEndDate As Label
    Friend WithEvents lblStartDate As Label
    Friend WithEvents dtpEndDate As DateTimePicker
    Friend WithEvents dtpStartDate As DateTimePicker
    Friend WithEvents lblDate As Label
    Friend WithEvents lblEmploymentType As Label
    Friend WithEvents cmbEmploymentType As SergeUtils.EasyCompletionComboBox
    Friend WithEvents lblLeaveType As Label
    Friend WithEvents cmbLeaveType As SergeUtils.EasyCompletionComboBox
    Friend WithEvents lblEmployeeName As Label
    Friend WithEvents cmbEmployeeName As SergeUtils.EasyCompletionComboBox
End Class
