import React, { useState } from 'react';
import { Scan, Image as ImageIcon, Loader2, CheckCircle2, ChevronRight, ChevronLeft } from 'lucide-react';
import { performOCR } from '../lib/ocr';

const BDTPhotoInspector = ({ images = [] }) => {
  const [currentIndex, setCurrentIndex] = useState(0);
  const [ocrResults, setOcrResults] = useState({});
  const [scanning, setScanning] = useState({});

  if (images.length === 0) return null;

  const currentImage = images[currentIndex];

  const handleScan = async (idx) => {
    if (ocrResults[idx] || scanning[idx]) return;
    
    setScanning(prev => ({ ...prev, [idx]: true }));
    try {
      const text = await performOCR(images[idx].base64);
      setOcrResults(prev => ({ ...prev, [idx]: text }));
    } catch (err) {
      setOcrResults(prev => ({ ...prev, [idx]: "Failed to read" }));
    } finally {
      setScanning(prev => ({ ...prev, [idx]: false }));
    }
  };

  return (
    <div className="bg-premium-900/40 border border-premium-100/10 rounded-3xl overflow-hidden mt-6">
      <div className="p-6 border-b border-premium-100/10 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-indigo-600/20 text-indigo-500 rounded-xl flex items-center justify-center">
            <ImageIcon size={20} />
          </div>
          <div>
            <h4 className="text-sm font-black uppercase tracking-widest">Photo Inspector</h4>
            <p className="text-[10px] font-bold text-premium-400 mt-1">
              Analyzing {images.length} embedded evidence photos
            </p>
          </div>
        </div>
        
        <div className="flex items-center gap-2">
          <button 
            onClick={() => setCurrentIndex(prev => Math.max(0, prev - 1))}
            className="p-2 hover:bg-premium-800 rounded-lg text-premium-400 disabled:opacity-30"
            disabled={currentIndex === 0}
          >
            <ChevronLeft size={20} />
          </button>
          <span className="text-xs font-black px-2">
            {currentIndex + 1} / {images.length}
          </span>
          <button 
            onClick={() => setCurrentIndex(prev => Math.min(images.length - 1, prev + 1))}
            className="p-2 hover:bg-premium-800 rounded-lg text-premium-400 disabled:opacity-30"
            disabled={currentIndex === images.length - 1}
          >
            <ChevronRight size={20} />
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2">
        <div className="p-6 bg-black flex items-center justify-center min-h-[300px]">
          <ImagePreview src={currentImage.base64} />
        </div>
        
        <div className="p-8 flex flex-col justify-center space-y-6">
          <div>
            <span className="px-3 py-1 bg-indigo-500/10 text-indigo-400 rounded-full text-[9px] font-black uppercase tracking-widest">
                Source: {currentImage.sheetName}
            </span>
            <h3 className="text-xl font-black mt-4 tracking-tighter">Automated Value Extraction</h3>
            <p className="text-premium-400 text-xs mt-2 leading-relaxed">
              Use the OCR engine to read numerical values from the photo evidence and compare them with the Excel entries.
            </p>
          </div>

          <div className="bg-premium-800/50 rounded-2xl p-5 border border-premium-100/5 min-h-[100px] flex flex-col justify-center">
            {scanning[currentIndex] ? (
              <div className="flex flex-col items-center py-4">
                <Loader2 className="animate-spin text-indigo-500 mb-3" size={24} />
                <span className="text-[10px] font-black uppercase tracking-widest text-premium-400">Reading Image Data...</span>
              </div>
            ) : ocrResults[currentIndex] ? (
              <div>
                <span className="text-[10px] font-black uppercase tracking-widest text-emerald-500 flex items-center gap-2 mb-2">
                  <CheckCircle2 size={12} /> Extraction Complete
                </span>
                <div className="text-xl font-black font-mono text-emerald-400 bg-emerald-500/5 p-3 rounded-xl border border-emerald-500/10">
                  {ocrResults[currentIndex]}
                </div>
              </div>
            ) : (
              <button 
                onClick={() => handleScan(currentIndex)}
                className="w-full py-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-3 transition-all"
              >
                <Scan size={18} />
                Extract Values from Photo
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

const ImagePreview = ({ src }) => {
    return (
        <div className="relative group">
            <img 
                src={src} 
                alt="Evidence" 
                className="max-w-full max-h-[400px] rounded-lg shadow-2xl transition-transform group-hover:scale-[1.02]"
            />
        </div>
    );
};

export default BDTPhotoInspector;
