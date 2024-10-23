#pragma once

//using namespace Microsoft::Office::Interop::Excel;
//using namespace Microsoft::Office::Interop::Word;

namespace Practicum {
    using namespace Microsoft::VisualBasic;
    using namespace MyDLL1;
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Сводка для MainForm
	/// </summary>
	public ref class MainForm : public System::Windows::Forms::Form
	{
	public:
		MainForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MainForm()
		{
			if (components)
			{
				delete components;
			}
		}

    protected:

    private: System::Windows::Forms::Button^ btnTask;
    private: System::Windows::Forms::Button^ btnExit;















    private: System::Windows::Forms::Label^ lblArray2;
    private: System::Windows::Forms::Label^ lblArray1;
    private: System::Windows::Forms::Label^ label1;
	private: System::Windows::Forms::TableLayoutPanel^ tableLayoutPanel2;
    private: System::Windows::Forms::TableLayoutPanel^ tableLayoutPanel1;




    private: System::Windows::Forms::DataGridView^ dataGridView1;
    private: System::Windows::Forms::DataGridView^ dataGridView3;

    private: System::Windows::Forms::Label^ lblSum;
    private: System::Windows::Forms::DataGridView^ dataGridView2;


    private: System::Windows::Forms::Label^ label2;
    private: System::Windows::Forms::Button^ btnShakerSort;

    private: System::Windows::Forms::TableLayoutPanel^ tableLayoutPanel3;
    private: System::Windows::Forms::Button^ btnInsert;
    private: System::Windows::Forms::Button^ btnGen;
    private: System::Windows::Forms::Button^ btnDelete;
    private: System::Windows::Forms::Button^ btnWord;
    private: System::Windows::Forms::Button^ btnExcel;














	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
            System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MainForm::typeid));
            System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle4 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
            System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle5 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
            System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle6 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
            this->btnTask = (gcnew System::Windows::Forms::Button());
            this->btnExit = (gcnew System::Windows::Forms::Button());
            this->lblArray2 = (gcnew System::Windows::Forms::Label());
            this->lblArray1 = (gcnew System::Windows::Forms::Label());
            this->label1 = (gcnew System::Windows::Forms::Label());
            this->tableLayoutPanel2 = (gcnew System::Windows::Forms::TableLayoutPanel());
            this->dataGridView3 = (gcnew System::Windows::Forms::DataGridView());
            this->label2 = (gcnew System::Windows::Forms::Label());
            this->dataGridView2 = (gcnew System::Windows::Forms::DataGridView());
            this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
            this->btnShakerSort = (gcnew System::Windows::Forms::Button());
            this->tableLayoutPanel1 = (gcnew System::Windows::Forms::TableLayoutPanel());
            this->btnExcel = (gcnew System::Windows::Forms::Button());
            this->lblSum = (gcnew System::Windows::Forms::Label());
            this->btnInsert = (gcnew System::Windows::Forms::Button());
            this->btnWord = (gcnew System::Windows::Forms::Button());
            this->btnDelete = (gcnew System::Windows::Forms::Button());
            this->tableLayoutPanel3 = (gcnew System::Windows::Forms::TableLayoutPanel());
            this->btnGen = (gcnew System::Windows::Forms::Button());
            this->tableLayoutPanel2->SuspendLayout();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView3))->BeginInit();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView2))->BeginInit();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
            this->tableLayoutPanel1->SuspendLayout();
            this->tableLayoutPanel3->SuspendLayout();
            this->SuspendLayout();
            // 
            // btnTask
            // 
            this->btnTask->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnTask->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnTask->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnTask->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnTask->Location = System::Drawing::Point(3, 131);
            this->btnTask->Name = L"btnTask";
            this->btnTask->Size = System::Drawing::Size(151, 66);
            this->btnTask->TabIndex = 32;
            this->btnTask->Text = L"Вычислить";
            this->btnTask->UseVisualStyleBackColor = false;
            this->btnTask->Click += gcnew System::EventHandler(this, &MainForm::btnTask_Click);
            // 
            // btnExit
            // 
            this->btnExit->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnExit->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnExit->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnExit->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnExit->Location = System::Drawing::Point(812, 13);
            this->btnExit->Name = L"btnExit";
            this->btnExit->Size = System::Drawing::Size(128, 67);
            this->btnExit->TabIndex = 33;
            this->btnExit->Text = L"Выход";
            this->btnExit->UseVisualStyleBackColor = false;
            this->btnExit->Click += gcnew System::EventHandler(this, &MainForm::btnExit_Click);
            // 
            // lblArray2
            // 
            this->lblArray2->Anchor = System::Windows::Forms::AnchorStyles::Left;
            this->lblArray2->AutoSize = true;
            this->lblArray2->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->lblArray2->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->lblArray2->Location = System::Drawing::Point(3, 124);
            this->lblArray2->Name = L"lblArray2";
            this->lblArray2->Size = System::Drawing::Size(225, 20);
            this->lblArray2->TabIndex = 29;
            this->lblArray2->Text = L"Результирующий массив";
            // 
            // lblArray1
            // 
            this->lblArray1->Anchor = System::Windows::Forms::AnchorStyles::Left;
            this->lblArray1->AutoSize = true;
            this->lblArray1->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->lblArray1->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->lblArray1->Location = System::Drawing::Point(3, 0);
            this->lblArray1->Name = L"lblArray1";
            this->lblArray1->Size = System::Drawing::Size(236, 20);
            this->lblArray1->TabIndex = 28;
            this->lblArray1->Text = L"Сформированный массив";
            this->lblArray1->Click += gcnew System::EventHandler(this, &MainForm::lblArray1_Click);
            // 
            // label1
            // 
            this->label1->Anchor = System::Windows::Forms::AnchorStyles::Top;
            this->label1->AutoSize = true;
            this->label1->BorderStyle = System::Windows::Forms::BorderStyle::FixedSingle;
            this->label1->Font = (gcnew System::Drawing::Font(L"Times New Roman", 14.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->label1->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->label1->Location = System::Drawing::Point(8, 9);
            this->label1->Name = L"label1";
            this->label1->Size = System::Drawing::Size(971, 65);
            this->label1->TabIndex = 59;
            this->label1->Text = resources->GetString(L"label1.Text");
            this->label1->TextAlign = System::Drawing::ContentAlignment::TopCenter;
            this->label1->Click += gcnew System::EventHandler(this, &MainForm::label1_Click);
            // 
            // tableLayoutPanel2
            // 
            this->tableLayoutPanel2->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
                | System::Windows::Forms::AnchorStyles::Left)
                | System::Windows::Forms::AnchorStyles::Right));
            this->tableLayoutPanel2->ColumnCount = 1;
            this->tableLayoutPanel2->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle()));
            this->tableLayoutPanel2->Controls->Add(this->lblArray2, 0, 3);
            this->tableLayoutPanel2->Controls->Add(this->lblArray1, 0, 0);
            this->tableLayoutPanel2->Controls->Add(this->dataGridView3, 0, 7);
            this->tableLayoutPanel2->Controls->Add(this->label2, 0, 6);
            this->tableLayoutPanel2->Controls->Add(this->dataGridView2, 0, 4);
            this->tableLayoutPanel2->Controls->Add(this->dataGridView1, 0, 1);
            this->tableLayoutPanel2->Location = System::Drawing::Point(8, 96);
            this->tableLayoutPanel2->Name = L"tableLayoutPanel2";
            this->tableLayoutPanel2->RowCount = 8;
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel2->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
            this->tableLayoutPanel2->Size = System::Drawing::Size(761, 354);
            this->tableLayoutPanel2->TabIndex = 64;
            // 
            // dataGridView3
            // 
            this->dataGridView3->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
                | System::Windows::Forms::AnchorStyles::Left)
                | System::Windows::Forms::AnchorStyles::Right));
            this->dataGridView3->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::AllCells;
            this->dataGridView3->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllCells;
            this->dataGridView3->BackgroundColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(17)),
                static_cast<System::Int32>(static_cast<System::Byte>(129)), static_cast<System::Int32>(static_cast<System::Byte>(178)));
            this->dataGridView3->ColumnHeadersVisible = false;
            dataGridViewCellStyle4->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleRight;
            dataGridViewCellStyle4->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(17)), static_cast<System::Int32>(static_cast<System::Byte>(129)),
                static_cast<System::Int32>(static_cast<System::Byte>(178)));
            dataGridViewCellStyle4->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            dataGridViewCellStyle4->ForeColor = System::Drawing::SystemColors::ActiveCaptionText;
            dataGridViewCellStyle4->Format = L"N0";
            dataGridViewCellStyle4->NullValue = nullptr;
            dataGridViewCellStyle4->SelectionBackColor = System::Drawing::SystemColors::Highlight;
            dataGridViewCellStyle4->SelectionForeColor = System::Drawing::SystemColors::ActiveCaptionText;
            dataGridViewCellStyle4->WrapMode = System::Windows::Forms::DataGridViewTriState::False;
            this->dataGridView3->DefaultCellStyle = dataGridViewCellStyle4;
            this->dataGridView3->Location = System::Drawing::Point(3, 271);
            this->dataGridView3->Name = L"dataGridView3";
            this->dataGridView3->ReadOnly = true;
            this->dataGridView3->RowHeadersVisible = false;
            this->dataGridView3->ScrollBars = System::Windows::Forms::ScrollBars::Horizontal;
            this->dataGridView3->Size = System::Drawing::Size(755, 80);
            this->dataGridView3->TabIndex = 0;
            this->dataGridView3->CellContentClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &MainForm::dataGridView2_CellContentClick);
            // 
            // label2
            // 
            this->label2->Anchor = System::Windows::Forms::AnchorStyles::Left;
            this->label2->AutoSize = true;
            this->label2->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->label2->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->label2->Location = System::Drawing::Point(3, 248);
            this->label2->Name = L"label2";
            this->label2->Size = System::Drawing::Size(389, 20);
            this->label2->TabIndex = 31;
            this->label2->Text = L"Отсортированный результирующий массив";
            // 
            // dataGridView2
            // 
            this->dataGridView2->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
                | System::Windows::Forms::AnchorStyles::Left)
                | System::Windows::Forms::AnchorStyles::Right));
            this->dataGridView2->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::AllCells;
            this->dataGridView2->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllCells;
            this->dataGridView2->BackgroundColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(17)),
                static_cast<System::Int32>(static_cast<System::Byte>(129)), static_cast<System::Int32>(static_cast<System::Byte>(178)));
            this->dataGridView2->ColumnHeadersVisible = false;
            dataGridViewCellStyle5->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleRight;
            dataGridViewCellStyle5->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(17)), static_cast<System::Int32>(static_cast<System::Byte>(129)),
                static_cast<System::Int32>(static_cast<System::Byte>(178)));
            dataGridViewCellStyle5->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            dataGridViewCellStyle5->ForeColor = System::Drawing::SystemColors::ActiveCaptionText;
            dataGridViewCellStyle5->Format = L"N0";
            dataGridViewCellStyle5->NullValue = nullptr;
            dataGridViewCellStyle5->SelectionBackColor = System::Drawing::SystemColors::Highlight;
            dataGridViewCellStyle5->SelectionForeColor = System::Drawing::SystemColors::ActiveCaptionText;
            dataGridViewCellStyle5->WrapMode = System::Windows::Forms::DataGridViewTriState::False;
            this->dataGridView2->DefaultCellStyle = dataGridViewCellStyle5;
            this->dataGridView2->Location = System::Drawing::Point(3, 147);
            this->dataGridView2->Name = L"dataGridView2";
            this->dataGridView2->ReadOnly = true;
            this->dataGridView2->RowHeadersVisible = false;
            this->dataGridView2->ScrollBars = System::Windows::Forms::ScrollBars::Horizontal;
            this->dataGridView2->Size = System::Drawing::Size(755, 78);
            this->dataGridView2->TabIndex = 30;
            // 
            // dataGridView1
            // 
            this->dataGridView1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
                | System::Windows::Forms::AnchorStyles::Left)
                | System::Windows::Forms::AnchorStyles::Right));
            this->dataGridView1->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::AllCells;
            this->dataGridView1->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllCells;
            this->dataGridView1->BackgroundColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(17)),
                static_cast<System::Int32>(static_cast<System::Byte>(129)), static_cast<System::Int32>(static_cast<System::Byte>(178)));
            this->dataGridView1->ColumnHeadersVisible = false;
            dataGridViewCellStyle6->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleRight;
            dataGridViewCellStyle6->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(17)), static_cast<System::Int32>(static_cast<System::Byte>(129)),
                static_cast<System::Int32>(static_cast<System::Byte>(178)));
            dataGridViewCellStyle6->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            dataGridViewCellStyle6->ForeColor = System::Drawing::SystemColors::ActiveCaptionText;
            dataGridViewCellStyle6->Format = L"N0";
            dataGridViewCellStyle6->NullValue = nullptr;
            dataGridViewCellStyle6->SelectionBackColor = System::Drawing::SystemColors::Highlight;
            dataGridViewCellStyle6->SelectionForeColor = System::Drawing::SystemColors::ActiveCaptionText;
            dataGridViewCellStyle6->WrapMode = System::Windows::Forms::DataGridViewTriState::False;
            this->dataGridView1->DefaultCellStyle = dataGridViewCellStyle6;
            this->dataGridView1->Location = System::Drawing::Point(3, 23);
            this->dataGridView1->Name = L"dataGridView1";
            this->dataGridView1->ReadOnly = true;
            this->dataGridView1->RowHeadersVisible = false;
            this->dataGridView1->ScrollBars = System::Windows::Forms::ScrollBars::Horizontal;
            this->dataGridView1->Size = System::Drawing::Size(755, 78);
            this->dataGridView1->TabIndex = 0;
            this->dataGridView1->CellContentClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &MainForm::dataGridView1_CellContentClick);
            // 
            // btnShakerSort
            // 
            this->btnShakerSort->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnShakerSort->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnShakerSort->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnShakerSort->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnShakerSort->Location = System::Drawing::Point(3, 255);
            this->btnShakerSort->Name = L"btnShakerSort";
            this->btnShakerSort->Size = System::Drawing::Size(152, 67);
            this->btnShakerSort->TabIndex = 66;
            this->btnShakerSort->Text = L"Шейкер сортировка";
            this->btnShakerSort->UseVisualStyleBackColor = false;
            this->btnShakerSort->Click += gcnew System::EventHandler(this, &MainForm::btnSort_Click);
            // 
            // tableLayoutPanel1
            // 
            this->tableLayoutPanel1->Anchor = System::Windows::Forms::AnchorStyles::Bottom;
            this->tableLayoutPanel1->ColumnCount = 6;
            this->tableLayoutPanel1->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Absolute,
                200)));
            this->tableLayoutPanel1->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Percent,
                20)));
            this->tableLayoutPanel1->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Percent,
                20)));
            this->tableLayoutPanel1->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Percent,
                20)));
            this->tableLayoutPanel1->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Percent,
                20)));
            this->tableLayoutPanel1->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Percent,
                20)));
            this->tableLayoutPanel1->Controls->Add(this->btnExcel, 4, 0);
            this->tableLayoutPanel1->Controls->Add(this->lblSum, 0, 0);
            this->tableLayoutPanel1->Controls->Add(this->btnExit, 5, 0);
            this->tableLayoutPanel1->Controls->Add(this->btnInsert, 1, 0);
            this->tableLayoutPanel1->Controls->Add(this->btnWord, 3, 0);
            this->tableLayoutPanel1->Controls->Add(this->btnDelete, 2, 0);
            this->tableLayoutPanel1->Location = System::Drawing::Point(15, 486);
            this->tableLayoutPanel1->Name = L"tableLayoutPanel1";
            this->tableLayoutPanel1->RowCount = 1;
            this->tableLayoutPanel1->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 100)));
            this->tableLayoutPanel1->Size = System::Drawing::Size(953, 94);
            this->tableLayoutPanel1->TabIndex = 65;
            this->tableLayoutPanel1->Paint += gcnew System::Windows::Forms::PaintEventHandler(this, &MainForm::tableLayoutPanel1_Paint);
            // 
            // btnExcel
            // 
            this->btnExcel->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnExcel->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnExcel->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnExcel->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnExcel->Location = System::Drawing::Point(662, 13);
            this->btnExcel->Name = L"btnExcel";
            this->btnExcel->Size = System::Drawing::Size(126, 67);
            this->btnExcel->TabIndex = 70;
            this->btnExcel->Text = L"Запись в Excel";
            this->btnExcel->UseVisualStyleBackColor = false;
            this->btnExcel->Click += gcnew System::EventHandler(this, &MainForm::btnExcel_Click);
            // 
            // lblSum
            // 
            this->lblSum->Anchor = System::Windows::Forms::AnchorStyles::Left;
            this->lblSum->AutoSize = true;
            this->lblSum->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->lblSum->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->lblSum->Location = System::Drawing::Point(3, 35);
            this->lblSum->Name = L"lblSum";
            this->lblSum->Size = System::Drawing::Size(159, 23);
            this->lblSum->TabIndex = 34;
            this->lblSum->Text = L"Сумма индексов:";
            this->lblSum->Visible = false;
            this->lblSum->Click += gcnew System::EventHandler(this, &MainForm::label2_Click);
            // 
            // btnInsert
            // 
            this->btnInsert->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnInsert->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnInsert->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnInsert->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnInsert->Location = System::Drawing::Point(212, 13);
            this->btnInsert->Name = L"btnInsert";
            this->btnInsert->Size = System::Drawing::Size(126, 67);
            this->btnInsert->TabIndex = 67;
            this->btnInsert->Text = L"Вставить элемент";
            this->btnInsert->UseVisualStyleBackColor = false;
            this->btnInsert->Click += gcnew System::EventHandler(this, &MainForm::btnInsert_Click);
            // 
            // btnWord
            // 
            this->btnWord->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnWord->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnWord->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnWord->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnWord->Location = System::Drawing::Point(512, 13);
            this->btnWord->Name = L"btnWord";
            this->btnWord->Size = System::Drawing::Size(126, 67);
            this->btnWord->TabIndex = 69;
            this->btnWord->Text = L"Запись в Word";
            this->btnWord->UseVisualStyleBackColor = false;
            this->btnWord->Click += gcnew System::EventHandler(this, &MainForm::btnWord_Click);
            // 
            // btnDelete
            // 
            this->btnDelete->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnDelete->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnDelete->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnDelete->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnDelete->Location = System::Drawing::Point(362, 13);
            this->btnDelete->Name = L"btnDelete";
            this->btnDelete->Size = System::Drawing::Size(126, 67);
            this->btnDelete->TabIndex = 68;
            this->btnDelete->Text = L"Удалить элемент";
            this->btnDelete->UseVisualStyleBackColor = false;
            this->btnDelete->Click += gcnew System::EventHandler(this, &MainForm::btnDelete_Click);
            // 
            // tableLayoutPanel3
            // 
            this->tableLayoutPanel3->Anchor = System::Windows::Forms::AnchorStyles::Right;
            this->tableLayoutPanel3->ColumnCount = 1;
            this->tableLayoutPanel3->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Percent,
                100)));
            this->tableLayoutPanel3->Controls->Add(this->btnGen, 0, 0);
            this->tableLayoutPanel3->Controls->Add(this->btnShakerSort, 0, 6);
            this->tableLayoutPanel3->Controls->Add(this->btnTask, 0, 3);
            this->tableLayoutPanel3->Location = System::Drawing::Point(797, 119);
            this->tableLayoutPanel3->Name = L"tableLayoutPanel3";
            this->tableLayoutPanel3->RowCount = 7;
            this->tableLayoutPanel3->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
            this->tableLayoutPanel3->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel3->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel3->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
            this->tableLayoutPanel3->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel3->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Absolute, 20)));
            this->tableLayoutPanel3->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
            this->tableLayoutPanel3->Size = System::Drawing::Size(158, 331);
            this->tableLayoutPanel3->TabIndex = 68;
            // 
            // btnGen
            // 
            this->btnGen->Anchor = System::Windows::Forms::AnchorStyles::None;
            this->btnGen->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(68)), static_cast<System::Int32>(static_cast<System::Byte>(68)),
                static_cast<System::Int32>(static_cast<System::Byte>(155)));
            this->btnGen->Font = (gcnew System::Drawing::Font(L"Times New Roman", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
                static_cast<System::Byte>(204)));
            this->btnGen->ForeColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(221)), static_cast<System::Int32>(static_cast<System::Byte>(237)),
                static_cast<System::Int32>(static_cast<System::Byte>(244)));
            this->btnGen->Location = System::Drawing::Point(3, 8);
            this->btnGen->Name = L"btnGen";
            this->btnGen->Size = System::Drawing::Size(151, 66);
            this->btnGen->TabIndex = 67;
            this->btnGen->Text = L"Генерация";
            this->btnGen->UseVisualStyleBackColor = false;
            this->btnGen->Click += gcnew System::EventHandler(this, &MainForm::btnGen_Click);
            // 
            // MainForm
            // 
            this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
            this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
            this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(33)), static_cast<System::Int32>(static_cast<System::Byte>(34)),
                static_cast<System::Int32>(static_cast<System::Byte>(33)));
            this->ClientSize = System::Drawing::Size(980, 592);
            this->Controls->Add(this->tableLayoutPanel3);
            this->Controls->Add(this->tableLayoutPanel1);
            this->Controls->Add(this->tableLayoutPanel2);
            this->Controls->Add(this->label1);
            this->MinimumSize = System::Drawing::Size(996, 631);
            this->Name = L"MainForm";
            this->Text = L"Практическая работа №1";
            this->Load += gcnew System::EventHandler(this, &MainForm::MainForm_Load);
            this->tableLayoutPanel2->ResumeLayout(false);
            this->tableLayoutPanel2->PerformLayout();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView3))->EndInit();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView2))->EndInit();
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
            this->tableLayoutPanel1->ResumeLayout(false);
            this->tableLayoutPanel1->PerformLayout();
            this->tableLayoutPanel3->ResumeLayout(false);
            this->ResumeLayout(false);
            this->PerformLayout();

        }
#pragma endregion
    private: System::Void tableLayoutPanel3_Paint(System::Object^ sender, System::Windows::Forms::PaintEventArgs^ e) {
    }
private: System::Void lblArray1_Click(System::Object^ sender, System::EventArgs^ e) {
}
private: System::Void rchTxtBox1_TextChanged(System::Object^ sender, System::EventArgs^ e) {
}
private: System::Void btnExit_Click(System::Object^ sender, System::EventArgs^ e) {
    Owner->Show();
    this->Close();
}
private: System::Void btnTask_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        int j = 0, k;
        int n = dataGridView1->ColumnCount;
        int* mas = new int[n];
        MyArrayOperation::ReadArray(mas, dataGridView1);
        int* newmas = new int[n];
        k = MyArrayOperation::sum(mas, n);
        j = MyArrayOperation::new_array(mas, newmas, n, k, j);
        if (j > 0) MyArrayOperation::output_array(newmas, j, dataGridView2);
        else MessageBox::Show("Полученный массив пуст", "", MessageBoxButtons::OK, MessageBoxIcon::Information);
        MessageBox::Show("Сумма индексов элементов, больше 30 и кратных 2: " + Convert::ToString(k), "", MessageBoxButtons::OK, MessageBoxIcon::Information);
        lblSum->Text = "Сумма индексов: " + Convert::ToString(k);
        lblSum->Visible = true;
        delete[] mas;
        delete[] newmas;
    }
    catch (Exception^ ex) {
		MessageBox::Show(ex->Message, "Ошибка ввода!", MessageBoxButtons::OK, MessageBoxIcon::Information);
		return;
	}
}
private: System::Void dataGridView2_CellContentClick(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
}

private: System::Void MainForm_Load(System::Object^ sender, System::EventArgs^ e) {
}
private: System::Void label2_Click(System::Object^ sender, System::EventArgs^ e) {
}
private: System::Void tableLayoutPanel1_Paint(System::Object^ sender, System::Windows::Forms::PaintEventArgs^ e) {
}
private: System::Void btnSort_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        int n = dataGridView2->ColumnCount;
        if (n == 0) {
            MessageBox::Show("Пустой массив нельзя отсортировать", "Ошибка сортировки!", MessageBoxButtons::OK, MessageBoxIcon::Error);
            return;
        }
        int* mas = new int[n];
        MyArrayOperation::ReadArray(mas, dataGridView2);
        MyArrayOperation::ShakerSort(mas, n);
        MyArrayOperation::output_array(mas, n, dataGridView3);
        delete[] mas;
    }
    catch (Exception^ ex) {
        MessageBox::Show(ex->Message, "Ошибка ввода!", MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }
}
private: System::Void btnInsert_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        dataGridView2->Rows->Clear();
        dataGridView2->Columns->Clear();
        dataGridView3->Rows->Clear();
        dataGridView3->Columns->Clear();
        lblSum->Visible = false;
        String^ g = Interaction::InputBox("Введите вставляемый элемент", "Окно ввода данных", "", -1, -1);
        int x = Convert::ToInt32(g);
        g = Interaction::InputBox("На какое место в массиве вы хотите его вставить элемент?", "Окно ввода данных", "", -1, -1);
        int k = Convert::ToInt32(g);
        int n = dataGridView1->ColumnCount;
        int* mas = new int[n + 1];
        MyArrayOperation::ReadArray(mas, dataGridView1);
        MyArrayOperation::InsertElem(mas, n, x, k);
        MyArrayOperation::output_array(mas, n + 1, dataGridView1);
        delete[] mas;
    }
    catch (Exception^ ex) {
        MessageBox::Show(ex->Message, "Ошибка ввода!", MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }
}
private: System::Void btnGen_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        dataGridView2->Rows->Clear();
        dataGridView2->Columns->Clear();
        dataGridView3->Rows->Clear();
        dataGridView3->Columns->Clear();
        lblSum->Visible = false;
        int k;
        String^ g = Interaction::InputBox("Введите количество элементов в массиве", "Ввод", "", -1, -1);
        int n = Convert::ToInt16(g);
        int* mas = new int[n];
        MyArrayOperation::enter_array(mas, n);
        MyArrayOperation::output_array(mas, n, dataGridView1);
        k = MyArrayOperation::sum(mas, n);
        delete[] mas;
    }
    catch (Exception^ ex) {
        MessageBox::Show(ex->Message, "Ошибка ввода!", MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }
}
private: System::Void label1_Click(System::Object^ sender, System::EventArgs^ e) {
}
private: System::Void btnDelete_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        dataGridView2->Rows->Clear();
        dataGridView2->Columns->Clear();
        dataGridView3->Rows->Clear();
        dataGridView3->Columns->Clear();
        lblSum->Visible = false;
        String^ g = Interaction::InputBox("Какой элемент вы хотите удалить?", "Окно ввода данных", "", -1, -1);
        int k = Convert::ToInt32(g);
        int n = dataGridView1->ColumnCount;
        int* mas = new int[n];
        MyArrayOperation::ReadArray(mas, dataGridView1);
        MyArrayOperation::DeleteElem(mas, n, k);
        MyArrayOperation::output_array(mas, n - 1, dataGridView1);
        delete[] mas;
    }
    catch (Exception^ ex) {
        MessageBox::Show(ex->Message, "Ошибка ввода!", MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }
}
private: System::Void btnWord_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        int n = dataGridView1->Columns->Count;
        int* mas = new int[n];
        MyArrayOperation::ReadArray(mas, dataGridView1);
        int j = dataGridView2->Columns->Count;
        if (n == 0 || j == 0) {
            MessageBox::Show("Нельзя записать пустой массив!", "Ошибка записи", MessageBoxButtons::OK, MessageBoxIcon::Error);
            return;
        }
        int* newmas = new int[j];
        MyArrayOperation::ReadArray(newmas, dataGridView2);
        MyArrayOperation::WriteWord(mas, newmas, n, j);
        delete[] mas;
        delete[] newmas;
    }
    catch (Exception^ ex) {
        MessageBox::Show(ex->Message, "Ошибка ввода!", MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }
}
private: System::Void btnExcel_Click(System::Object^ sender, System::EventArgs^ e) {
    try {
        int n = dataGridView1->Columns->Count;
        int* mas = new int[n];
        MyArrayOperation::ReadArray(mas, dataGridView1);
        int j = dataGridView2->Columns->Count;
        if (n == 0 || j == 0) {
            MessageBox::Show("Нельзя записать пустой массив!", "Ошибка записи", MessageBoxButtons::OK, MessageBoxIcon::Error);
            return;
        }
        int* newmas = new int[j];
        MyArrayOperation::ReadArray(newmas, dataGridView2);
        MyArrayOperation::WriteExcel(mas, newmas, n, j);
        delete[] mas;
        delete[] newmas;
    }
    catch (Exception^ ex) {
        MessageBox::Show(ex->Message, "Ошибка ввода!", MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }
}
private: System::Void dataGridView1_CellContentClick(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
}
};
}
