import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Download, Upload, Users, FileSpreadsheet, Phone, Tag } from 'lucide-react';

const ExcelToVCFConverter = () => {
  const [file, setFile] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheets, setSelectedSheets] = useState(new Set());
  const [columnMapping, setColumnMapping] = useState({});
  const [headerRows, setHeaderRows] = useState({});
  const [useSheetAsTag, setUseSheetAsTag] = useState(true);
  const [processing, setProcessing] = useState(false);
  const [contacts, setContacts] = useState([]);

  const handleFileUpload = async (event) => {
    const uploadedFile = event.target.files[0];
    if (!uploadedFile) return;

    setFile(uploadedFile);
    setProcessing(true);

    try {
      const arrayBuffer = await uploadedFile.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      
      const sheetData = workbook.SheetNames.map(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        return {
          name: sheetName,
          data,
          totalRows: data.length
        };
      });

      setSheets(sheetData);
      setSelectedSheets(new Set(sheetData.map(s => s.name)));
      
      // Initialize header rows (default to row 1)
      const headerRowsInit = {};
      sheetData.forEach(sheet => {
        headerRowsInit[sheet.name] = 0; // 0-based index
      });
      setHeaderRows(headerRowsInit);
      
      // Auto-detect column mappings
      const mapping = {};
      sheetData.forEach(sheet => {
        const headers = sheet.data[0] || [];
        mapping[sheet.name] = detectColumnMapping(headers);
      });
      setColumnMapping(mapping);
      
    } catch (error) {
      alert('Erreur lors de la lecture du fichier Excel : ' + error.message);
    } finally {
      setProcessing(false);
    }
  };

  const detectColumnMapping = (headers) => {
    const mapping = {};
    
    headers.forEach((header, index) => {
      const lowerHeader = header?.toString().toLowerCase() || '';
      
      if (lowerHeader.includes('prénom') || (lowerHeader.includes('first') && lowerHeader.includes('name'))) {
        mapping.firstName = index;
      } else if (lowerHeader.includes('nom') || (lowerHeader.includes('last') && lowerHeader.includes('name'))) {
        mapping.lastName = index;
      } else if (lowerHeader.includes('nom complet') || (lowerHeader.includes('full') && lowerHeader.includes('name'))) {
        mapping.fullName = index;
      } else if (lowerHeader.includes('nom') && !mapping.firstName && !mapping.lastName) {
        mapping.fullName = index;
      } else if (lowerHeader.includes('téléphone') || lowerHeader.includes('phone') || lowerHeader.includes('mobile') || lowerHeader.includes('cell')) {
        if (!mapping.phone1) mapping.phone1 = index;
        else if (!mapping.phone2) mapping.phone2 = index;
        else if (!mapping.phone3) mapping.phone3 = index;
      } else if (lowerHeader.includes('email') || lowerHeader.includes('mail') || lowerHeader.includes('courriel')) {
        mapping.email = index;
      } else if (lowerHeader.includes('entreprise') || lowerHeader.includes('société') || lowerHeader.includes('company') || lowerHeader.includes('organization')) {
        mapping.organization = index;
      } else if (lowerHeader.includes('titre') || lowerHeader.includes('poste') || lowerHeader.includes('title') || lowerHeader.includes('job')) {
        mapping.title = index;
      } else if (lowerHeader.includes('adresse') || lowerHeader.includes('address')) {
        mapping.address = index;
      }
    });
    
    return mapping;
  };

  const updateHeaderRow = (sheetName, rowIndex) => {
    setHeaderRows(prev => ({
      ...prev,
      [sheetName]: rowIndex
    }));
    
    // Update column mapping based on new header row
    const sheet = sheets.find(s => s.name === sheetName);
    if (sheet && sheet.data[rowIndex]) {
      const headers = sheet.data[rowIndex];
      setColumnMapping(prev => ({
        ...prev,
        [sheetName]: detectColumnMapping(headers)
      }));
    }
  };

  const updateColumnMapping = (sheetName, field, columnIndex) => {
    setColumnMapping(prev => ({
      ...prev,
      [sheetName]: {
        ...prev[sheetName],
        [field]: columnIndex === -1 ? undefined : columnIndex
      }
    }));
  };

  const generateVCF = () => {
    const allContacts = [];
    
    sheets.forEach(sheet => {
      if (!selectedSheets.has(sheet.name)) return;
      
      const mapping = columnMapping[sheet.name] || {};
      const headerRowIndex = headerRows[sheet.name] || 0;
      const dataRows = sheet.data.slice(headerRowIndex + 1);
      
      dataRows.forEach(row => {
        if (!row || row.every(cell => !cell)) return; // Skip empty rows
        
        const contact = {};
        
        // Name handling
        if (mapping.fullName !== undefined) {
          contact.fullName = row[mapping.fullName] || '';
        } else {
          const firstName = row[mapping.firstName] || '';
          const lastName = row[mapping.lastName] || '';
          contact.fullName = `${firstName} ${lastName}`.trim();
        }
        
        // Multiple phone fields
        contact.phones = [];
        if (mapping.phone1 !== undefined && row[mapping.phone1]) {
          contact.phones.push({ type: 'TEL;TYPE=CELL', value: row[mapping.phone1] });
        }
        if (mapping.phone2 !== undefined && row[mapping.phone2]) {
          contact.phones.push({ type: 'TEL;TYPE=WORK', value: row[mapping.phone2] });
        }
        if (mapping.phone3 !== undefined && row[mapping.phone3]) {
          contact.phones.push({ type: 'TEL;TYPE=HOME', value: row[mapping.phone3] });
        }
        
        // Other fields
        contact.email = row[mapping.email] || '';
        contact.organization = row[mapping.organization] || '';
        contact.title = row[mapping.title] || '';
        contact.address = row[mapping.address] || '';
        
        // Add sheet name as tag if enabled
        if (useSheetAsTag) {
          contact.categories = sheet.name;
        }
        
        if (contact.fullName || contact.phones.length > 0 || contact.email) {
          allContacts.push(contact);
        }
      });
    });
    
    setContacts(allContacts);
    
    // Generate VCF content
    let vcfContent = '';
    
    allContacts.forEach(contact => {
      vcfContent += 'BEGIN:VCARD\n';
      vcfContent += 'VERSION:3.0\n';
      
      if (contact.fullName) {
        vcfContent += `FN:${contact.fullName}\n`;
        vcfContent += `N:${contact.fullName};;;;\n`;
      }
      
      // Multiple phone numbers
      contact.phones.forEach(phone => {
        vcfContent += `${phone.type}:${phone.value}\n`;
      });
      
      if (contact.email) {
        vcfContent += `EMAIL:${contact.email}\n`;
      }
      
      if (contact.organization) {
        vcfContent += `ORG:${contact.organization}\n`;
      }
      
      if (contact.title) {
        vcfContent += `TITLE:${contact.title}\n`;
      }
      
      if (contact.address) {
        vcfContent += `ADR:;;${contact.address};;;;\n`;
      }
      
      if (contact.categories) {
        vcfContent += `CATEGORIES:${contact.categories}\n`;
      }
      
      vcfContent += 'END:VCARD\n\n';
    });
    
    // Try multiple download methods
    try {
      // Method 1: Modern approach with better browser support
      const blob = new Blob([vcfContent], { type: 'text/vcard;charset=utf-8' });
      
      // Check if we can use the modern download API
      if (navigator.msSaveBlob) {
        // IE10+
        navigator.msSaveBlob(blob, `contacts_${new Date().toISOString().split('T')[0]}.vcf`);
      } else {
        // Modern browsers
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `contacts_${new Date().toISOString().split('T')[0]}.vcf`;
        link.style.display = 'none';
        
        // Ensure the link is added to DOM before clicking
        document.body.appendChild(link);
        
        // Force download
        setTimeout(() => {
          link.click();
          
          // Clean up after a delay
          setTimeout(() => {
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
          }, 100);
        }, 100);
      }
    } catch (error) {
      console.error('Download failed:', error);
      // Fallback: show content in a new window for manual save
      const newWindow = window.open('', '_blank');
      newWindow.document.write(`
        <html>
          <head>
            <title>Fichier VCF - Contacts</title>
            <style>
              body { font-family: monospace; padding: 20px; }
              .instructions { background: #f0f0f0; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
              .vcf-content { white-space: pre-wrap; background: #f9f9f9; padding: 15px; border: 1px solid #ddd; }
            </style>
          </head>
          <body>
            <div class="instructions">
              <h3>Instructions de téléchargement manuel :</h3>
              <p>1. Sélectionnez tout le contenu ci-dessous (Ctrl+A)</p>
              <p>2. Copiez le contenu (Ctrl+C)</p>
              <p>3. Ouvrez un éditeur de texte (Bloc-notes, etc.)</p>
              <p>4. Collez le contenu (Ctrl+V)</p>
              <p>5. Enregistrez le fichier avec l'extension .vcf</p>
            </div>
            <div class="vcf-content">${vcfContent}</div>
          </body>
        </html>
      `);
    }
  };

  return (
    <div className="max-w-6xl mx-auto p-6 bg-gradient-to-br from-blue-50 to-indigo-100 min-h-screen">
      <div className="bg-white rounded-xl shadow-xl p-8">
        <div className="text-center mb-8">
          <div className="flex justify-center items-center gap-3 mb-4">
            <FileSpreadsheet className="w-8 h-8 text-green-600" />
            <span className="text-2xl font-bold text-gray-800">→</span>
            <Phone className="w-8 h-8 text-blue-600" />
          </div>
          <h1 className="text-3xl font-bold text-gray-800 mb-2">Convertisseur Excel vers VCF</h1>
          <p className="text-gray-600">Convertissez vos feuilles de contacts Excel au format VCF pour votre téléphone</p>
        </div>

        {/* File Upload */}
        <div className="mb-8">
          <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-blue-300 rounded-lg cursor-pointer bg-blue-50 hover:bg-blue-100 transition-colors">
            <div className="flex flex-col items-center justify-center pt-5 pb-6">
              <Upload className="w-8 h-8 mb-4 text-blue-500" />
              <p className="mb-2 text-sm text-gray-500">
                <span className="font-semibold">Cliquez pour télécharger</span> votre fichier Excel
              </p>
              <p className="text-xs text-gray-500">Fichiers XLSX, XLS supportés</p>
            </div>
            <input
              type="file"
              className="hidden"
              accept=".xlsx,.xls"
              onChange={handleFileUpload}
            />
          </label>
        </div>

        {processing && (
          <div className="text-center py-8">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
            <p className="mt-2 text-gray-600">Traitement du fichier Excel...</p>
          </div>
        )}

        {sheets.length > 0 && !processing && (
          <>
            {/* Tag Option */}
            <div className="mb-8 p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
              <label className="flex items-center gap-3 cursor-pointer">
                <input
                  type="checkbox"
                  checked={useSheetAsTag}
                  onChange={(e) => setUseSheetAsTag(e.target.checked)}
                  className="w-4 h-4 text-blue-600"
                />
                <Tag className="w-5 h-5 text-yellow-600" />
                <div>
                  <div className="font-medium text-gray-800">Utiliser les noms de feuilles comme étiquettes</div>
                  <div className="text-sm text-gray-600">Ajoute le nom de la feuille comme catégorie pour chaque contact</div>
                </div>
              </label>
            </div>

            {/* Sheet Selection */}
            <div className="mb-8">
              <div className="mb-4">
                <h2 className="text-xl font-semibold mb-3 flex items-center gap-2">
                  <Users className="w-5 h-5" />
                  Sélectionner les feuilles à convertir
                </h2>
                <div className="flex flex-col sm:flex-row gap-2">
                  <button
                    onClick={() => setSelectedSheets(new Set(sheets.map(s => s.name)))}
                    className="flex-1 px-3 py-2 text-sm bg-blue-100 hover:bg-blue-200 text-blue-700 rounded-md transition-colors text-center"
                  >
                    Tout sélectionner
                  </button>
                  <button
                    onClick={() => setSelectedSheets(new Set())}
                    className="flex-1 px-3 py-2 text-sm bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-md transition-colors text-center"
                  >
                    Tout désélectionner
                  </button>
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {sheets.map(sheet => (
                  <div key={sheet.name} className="border rounded-lg p-4">
                    <label className="flex items-center gap-3 cursor-pointer mb-3">
                      <input
                        type="checkbox"
                        checked={selectedSheets.has(sheet.name)}
                        onChange={(e) => {
                          const newSelected = new Set(selectedSheets);
                          if (e.target.checked) {
                            newSelected.add(sheet.name);
                          } else {
                            newSelected.delete(sheet.name);
                          }
                          setSelectedSheets(newSelected);
                        }}
                        className="w-4 h-4 text-blue-600"
                      />
                      <div>
                        <div className="font-medium text-gray-800">{sheet.name}</div>
                        <div className="text-sm text-gray-500">{sheet.totalRows} lignes</div>
                      </div>
                    </label>
                    
                    {selectedSheets.has(sheet.name) && (
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          Ligne d'en-têtes :
                        </label>
                        <select
                          value={headerRows[sheet.name] || 0}
                          onChange={(e) => updateHeaderRow(sheet.name, parseInt(e.target.value))}
                          className="w-full p-2 border border-gray-300 rounded-md text-sm"
                        >
                          {sheet.data.slice(0, Math.min(10, sheet.data.length)).map((row, index) => (
                            <option key={index} value={index}>
                              Ligne {index + 1}: {row.slice(0, 3).join(', ')}...
                            </option>
                          ))}
                        </select>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>

            {/* Column Mapping */}
            {Array.from(selectedSheets).map(sheetName => {
              const sheet = sheets.find(s => s.name === sheetName);
              const mapping = columnMapping[sheetName] || {};
              const headerRowIndex = headerRows[sheetName] || 0;
              const headers = sheet.data[headerRowIndex] || [];
              
              return (
                <div key={sheetName} className="mb-8 border rounded-lg p-6 bg-gray-50">
                  <h3 className="text-lg font-semibold mb-4">Mappage des colonnes pour "{sheetName}"</h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    {[
                      { key: 'fullName', label: 'Nom complet' },
                      { key: 'firstName', label: 'Prénom' },
                      { key: 'lastName', label: 'Nom de famille' },
                      { key: 'phone1', label: 'Téléphone 1' },
                      { key: 'phone2', label: 'Téléphone 2 (Travail)' },
                      { key: 'phone3', label: 'Téléphone 3 (Domicile)' },
                      { key: 'email', label: 'Email' },
                      { key: 'organization', label: 'Organisation' },
                      { key: 'title', label: 'Titre/Poste' },
                      { key: 'address', label: 'Adresse' }
                    ].map(field => (
                      <div key={field.key}>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          {field.label}
                        </label>
                        <select
                          value={mapping[field.key] ?? -1}
                          onChange={(e) => updateColumnMapping(sheetName, field.key, parseInt(e.target.value))}
                          className="w-full p-2 border border-gray-300 rounded-md text-sm"
                        >
                          <option value={-1}>-- Non mappé --</option>
                          {headers.map((header, index) => (
                            <option key={index} value={index}>
                              {header || `Colonne ${index + 1}`}
                            </option>
                          ))}
                        </select>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}

            {/* Generate Button */}
            <div className="text-center">
              <button
                onClick={generateVCF}
                disabled={selectedSheets.size === 0}
                className="bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white px-8 py-3 rounded-lg font-semibold flex items-center gap-2 mx-auto transition-colors"
              >
                <Download className="w-5 h-5" />
                Générer le fichier VCF
              </button>
              {selectedSheets.size === 0 && (
                <p className="text-sm text-gray-500 mt-2">Veuillez sélectionner au moins une feuille</p>
              )}
            </div>
          </>
        )}

        {contacts.length > 0 && (
          <div className="mt-8 p-4 bg-green-50 border border-green-200 rounded-lg">
            <div className="flex items-center gap-2 text-green-800">
              <Users className="w-5 h-5" />
              <span className="font-semibold">Succès !</span>
            </div>
            <p className="text-green-700 mt-1">
              Fichier VCF généré avec {contacts.length} contacts. Le fichier a été téléchargé sur votre appareil.
            </p>
          </div>
        )}

        <div className="mt-8 p-4 bg-blue-50 border border-blue-200 rounded-lg">
          <h3 className="font-semibold text-blue-800 mb-2">Comment utiliser le fichier VCF :</h3>
          <div className="text-sm text-blue-700 space-y-1">
            <p>• <strong>iPhone :</strong> Envoyez-vous le fichier VCF par email, ouvrez-le dans Mail, et appuyez sur "Ajouter tous les contacts"</p>
            <p>• <strong>Android :</strong> Copiez le fichier VCF sur votre téléphone et ouvrez-le avec l'application Contacts</p>
            <p>• <strong>Google Contacts :</strong> Allez sur contacts.google.com → Importer → Téléchargez le fichier VCF</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ExcelToVCFConverter;
