const FORM_ID = 'ISI_ID_FORMULIR_TUJUAN'; 

/**
 * Fungsi untuk update Form yang sudah ada, menggunakan penghapusan item manual 
 * untuk menghindari TypeError: deleteAllItems.
 */
function updateExistingFormReliable() { // Ganti nama fungsi agar berbeda
  Logger.log("--- Mulai Update Form ---");
  
  // --- 1. MENGAMBIL DATA DARI SHEET ---
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Sheet1');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet bernama 'Sheet1' tidak ditemukan! Harap periksa nama sheet.");
    return;
  }
  
  var numberRows = sheet.getDataRange().getNumRows();
  Logger.log("Jumlah baris data yang ditemukan: " + numberRows);
  
  if (numberRows <= 0) { 
    SpreadsheetApp.getUi().alert("Sheet tidak memiliki data yang cukup untuk dibuat kuis.");
    return;
  }

  // Pengambilan data
  var myQuestions = sheet.getRange(1, 1, numberRows, 1).getValues(); 
  var myAnswers = sheet.getRange(1, 2, numberRows, 1).getValues();   
  var myGuesses = sheet.getRange(1, 2, numberRows, 5).getValues();   

  var myShuffled = myGuesses.map(shuffleEachRow);
  
  // --- 2. MEMBUKA DAN MEMBERSIHKAN FORMULIR ---
  try {
    var form = FormApp.openById(FORM_ID);
    Logger.log("Formulir berhasil dibuka. Judul: " + form.getTitle());
    
    // **METODE PENGHAPUSAN ITEM MANUAL**
    Logger.log("Mencoba menghapus semua item lama secara manual...");
    var items = form.getItems();
    // Hapus dari indeks terakhir (untuk menghindari perubahan indeks saat menghapus)
    for (var i = items.length - 1; i >= 0; i--) { 
      form.deleteItem(items[i]);
    }
    Logger.log("Semua item lama berhasil dihapus.");
    
    form.setIsQuiz(true); 
    
  } catch (e) {
    Logger.log("ERROR KRITIS saat membuka/menghapus item Form: " + e.toString());
    SpreadsheetApp.getUi().alert("ERROR KRITIS: " + e.toString());
    return;
  }

  // --- 3. MENAMBAHKAN ITEM BARU KE FORMULIR ---
  for (var i = 0; i < numberRows; i++) {
    var correctAnswerValue = myAnswers[i][0]; 
    var shuffledOptions = myShuffled[i];      
    
    var correctIndex = shuffledOptions.indexOf(correctAnswerValue);
    
    var addItem = form.addMultipleChoiceItem();
    var choices = [];
    
    for (var j = 0; j < shuffledOptions.length; j++) {
      var isCorrect = (j === correctIndex); 
      choices.push(
        addItem.createChoice(shuffledOptions[j], isCorrect)
      );
    }
    
    addItem.setTitle(myQuestions[i][0])
           .setPoints(1)
           .setChoices(choices);
  }
  
  Logger.log("--- Update Formulir Selesai ---");
  SpreadsheetApp.getUi().alert("Formulir dengan ID: " + FORM_ID + " telah berhasil diperbarui!");
}

function shuffleEachRow(array) {
  var i, j, temp;
  for (i = array.length - 1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}
