// =======================================================================
// KONFIGURASI: Ganti dengan ID Formulir Google Anda yang sebenarnya
// =======================================================================
const FORM_ID = '1HRQwhLPbuIY3QxRqLgbV1nCOPIcCp6Ye9tSmxyEwjjU'; 

/**
 * Fungsi utama untuk membaca data dari Google Sheet (Sheet1), 
 * mengacak pilihan, dan MENGUPDATE Form Google yang sudah ada.
 */
function updateFormWithMixedTypes() { 
  Logger.log("--- Mulai Update Form dengan Jenis Campuran ---");
  
  // --- 1. MENGAMBIL DATA DARI SHEET ---
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Sheet1');
  
  if (!sheet) {
    // Diganti alert menjadi Logger.log
    Logger.log("ERROR: Sheet bernama 'Sheet1' tidak ditemukan! Harap periksa nama sheet.");
    return;
  }
  
  var numberRows = sheet.getLastRow();
  if (numberRows <= 1) { 
    // Diganti alert menjadi Logger.log
    Logger.log("ERROR: Sheet tidak memiliki data yang cukup (minimal 2 baris data).");
    return;
  }

  // Mengambil 7 Kolom (Kolom A-G)
  var allData = sheet.getRange(1, 1, numberRows, 7).getValues(); 

  // --- 2. MEMBUKA DAN MEMBERSIHKAN FORMULIR ---
  try {
    var form = FormApp.openById(FORM_ID);
    
    // Metode Penghapusan Item Manual yang andal
    var items = form.getItems();
    for (var i = items.length - 1; i >= 0; i--) { 
      form.deleteItem(items[i]);
    }
    Logger.log("Semua item lama berhasil dihapus.");
    
    form.setIsQuiz(true); 
    
  } catch (e) {
    // Diganti alert menjadi Logger.log
    Logger.log("ERROR KRITIS saat membuka/menghapus item Form: " + e.toString() + ". Pastikan ID dan izin benar.");
    return;
  }

  // --- 3. ITERASI DAN MENAMBAHKAN ITEM BARU ---
  for (var i = 0; i < numberRows; i++) {
    var row = allData[i];
    var questionType = row[0].toString().toUpperCase().trim(); // Kolom A

    // PERBAIKAN KRITIS: Bersihkan dan validasi Judul di sini
    var cleanTitle = row[1] ? row[1].toString().trim() : "";  // Kolom B

    if (cleanTitle === "") {
        Logger.log("Baris " + (i + 1) + " dilewati karena judul pertanyaan kosong.");
        continue; 
    }
    
    // Ganti nilai di row[1] dengan versi yang sudah bersih (cleanTitle)
    row[1] = cleanTitle; 

    switch (questionType) {
      case 'PG':
        addMultipleChoiceItem(form, row);
        break;
      case 'DD': 
        addDropdownItem(form, row);
        break;
      case 'IS':
        addShortAnswerItem(form, row);
        break;
      case 'ESAI':
        addParagraphItem(form, row); 
        break;
      case 'NAMA': 
        addNameDropdown(form, row);
        break;
      default:
        Logger.log("Jenis pertanyaan tidak dikenal di baris " + (i + 1) + ": " + questionType + ". Dilewati.");
    }
  }
  
  Logger.log("--- Update Formulir Selesai ---");
  // Diganti alert menjadi Logger.log
  Logger.log("Formulir dengan ID: " + FORM_ID + " telah berhasil diperbarui!");
}


// =======================================================================
// FUNGSI PEMBANTU (HELPER FUNCTIONS)
// =======================================================================

/**
 * Menambahkan item Pilihan Ganda (Multiple Choice). (PG)
 */
function addMultipleChoiceItem(form, row) {
  var questionTitle = row[1]; 
  var myAnswers = row[2]; 
  var myGuesses = row.slice(2, 7); 

  var shuffledOptions = shuffleArray(myGuesses);
  
  var addItem = form.addMultipleChoiceItem();
  var choices = createChoices(addItem, shuffledOptions, myAnswers);
  
  addItem.setTitle(questionTitle)
         .setPoints(1)
         .setChoices(choices);
}

/**
 * Menambahkan item Dropdown (List). (DD)
 */
function addDropdownItem(form, row) {
  var questionTitle = row[1]; 
  var myAnswers = row[2]; 
  var myGuesses = row.slice(2, 7); 

  var shuffledOptions = shuffleArray(myGuesses);
  
  var addItem = form.addListItem();
  var choices = createChoices(addItem, shuffledOptions, myAnswers);

  addItem.setTitle(questionTitle)
         .setPoints(1)
         .setChoices(choices);
}

/**
 * Menambahkan item Isian Singkat (Short Answer/Text). (IS)
 */
function addShortAnswerItem(form, row) {
  var questionTitle = row[1]; 
  var correctAnswer = row[2].toString().trim();

  var addItem = form.addTextItem();
  addItem.setTitle(questionTitle)
         .setPoints(1); 
         
  if (correctAnswer !== "") {
    addItem.setValidation(
      FormApp.createTextValidation()
        .requireTextIsEqualTo(correctAnswer)
        .build()
    );
    var feedback = FormApp.createFeedback().setText('Jawaban yang benar adalah: ' + correctAnswer).build();
    addItem.setCorrectFeedback(feedback)
           .setIncorrectFeedback(feedback);
  }
}

/**
 * Menambahkan item Paragraf (Paragraph/Esai). (ESAI)
 */
function addParagraphItem(form, row) {
  var questionTitle = row[1]; 
  
  var addItem = form.addParagraphTextItem();
  addItem.setTitle(questionTitle);
}

/**
 * Menambahkan item Dropdown untuk Pilihan Nama. (NAMA)
 * Poin 0 dan tanpa Kunci Jawaban.
 */
function addNameDropdown(form, row) {
  var questionTitle = row[1];
  // Ambil pilihan dari Kolom C sampai G
  var nameOptions = row.slice(2, 7); 
  
  var addItem = form.addListItem();
  var choices = [];

  for (var j = 0; j < nameOptions.length; j++) {
    var optionValue = nameOptions[j].toString().trim();
    
    if (optionValue !== "") {
        // Tidak ada parameter 'isCorrect' (true/false) = Tidak ada kunci jawaban
        choices.push(
            addItem.createChoice(optionValue)
        );
    }
  }
  
  addItem.setTitle(questionTitle)
         .setPoints(0) // Poin disetel 0
         .setChoices(choices)
         .setRequired(true); 
}

/**
 * Fungsi untuk membuat array Choices untuk PG atau DD.
 */
function createChoices(item, shuffledOptions, myAnswers) {
    var choices = [];
    var correctIndex = shuffledOptions.indexOf(myAnswers);

    for (var j = 0; j < shuffledOptions.length; j++) {
        var isCorrect = (j === correctIndex);
        var optionValue = shuffledOptions[j].toString().trim();
        
        if (optionValue !== "") {
            choices.push(
                item.createChoice(optionValue, isCorrect)
            );
        }
    }
    return choices;
}

/**
 * Mengacak elemen dalam array (Algoritma Fisher-Yates).
 */
function shuffleArray(array) {
  var i, j, temp;
  for (i = array.length - 1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}
