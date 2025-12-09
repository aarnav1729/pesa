import { useState, useCallback } from 'react';
import { Upload, FileSpreadsheet, X, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { cn } from '@/lib/utils';

interface FileUploadProps {
  onFilesSelected: (files: File[]) => void;
  isProcessing: boolean;
}

export function FileUpload({ onFilesSelected, isProcessing }: FileUploadProps) {
  const [dragActive, setDragActive] = useState(false);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);

  const handleDrag = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    const files = Array.from(e.dataTransfer.files).filter(
      (file) => file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );
    
    if (files.length > 0) {
      setSelectedFiles((prev) => [...prev, ...files]);
    }
  }, []);

  const handleFileInput = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const files = Array.from(e.target.files);
      setSelectedFiles((prev) => [...prev, ...files]);
    }
  }, []);

  const removeFile = (index: number) => {
    setSelectedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const handleProcess = () => {
    if (selectedFiles.length > 0) {
      onFilesSelected(selectedFiles);
    }
  };

  return (
    <div className="w-full max-w-2xl mx-auto animate-fade-in">
      {/* Upload Area */}
      <div
        className={cn(
          'relative rounded-xl border-2 border-dashed p-12 transition-all duration-200 cursor-pointer',
          'bg-card shadow-enterprise',
          dragActive
            ? 'border-primary bg-primary/5 shadow-enterprise-lg scale-[1.01]'
            : 'border-border hover:border-primary/50 hover:shadow-enterprise-md'
        )}
        onDragEnter={handleDrag}
        onDragLeave={handleDrag}
        onDragOver={handleDrag}
        onDrop={handleDrop}
        onClick={() => document.getElementById('file-input')?.click()}
      >
        <input
          id="file-input"
          type="file"
          multiple
          accept=".xlsx,.xls"
          className="hidden"
          onChange={handleFileInput}
        />
        
        <div className="flex flex-col items-center gap-4 text-center">
          <div className={cn(
            'p-4 rounded-xl transition-all duration-200',
            dragActive ? 'gradient-primary scale-105' : 'bg-muted'
          )}>
            <Upload className={cn(
              'w-8 h-8 transition-colors',
              dragActive ? 'text-primary-foreground' : 'text-primary'
            )} />
          </div>
          
          <div>
            <h3 className="text-lg font-semibold text-foreground mb-1">
              Drop your XLSX files here
            </h3>
            <p className="text-sm text-muted-foreground">
              or click to browse • Multiple files supported
            </p>
          </div>
        </div>
      </div>

      {/* Selected Files */}
      {selectedFiles.length > 0 && (
        <div className="mt-6 space-y-3 animate-slide-up">
          <h4 className="text-sm font-medium text-muted-foreground">
            Selected Files ({selectedFiles.length})
          </h4>
          
          <div className="space-y-2 max-h-48 overflow-y-auto pr-2">
            {selectedFiles.map((file, index) => (
              <div
                key={`${file.name}-${index}`}
                className="flex items-center justify-between p-3 rounded-lg bg-card border border-border group hover:border-primary/30 transition-colors shadow-enterprise"
              >
                <div className="flex items-center gap-3">
                  <div className="p-2 rounded-md bg-primary/10">
                    <FileSpreadsheet className="w-4 h-4 text-primary" />
                  </div>
                  <span className="text-sm font-medium text-foreground truncate max-w-[280px]">
                    {file.name}
                  </span>
                  <span className="text-xs text-muted-foreground px-2 py-0.5 bg-muted rounded-md">
                    {(file.size / 1024).toFixed(1)} KB
                  </span>
                </div>
                
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    removeFile(index);
                  }}
                  className="p-1.5 rounded-md opacity-0 group-hover:opacity-100 hover:bg-destructive/10 transition-all"
                >
                  <X className="w-4 h-4 text-destructive" />
                </button>
              </div>
            ))}
          </div>

          <Button
            variant="gradient"
            size="lg"
            className="w-full mt-4"
            onClick={handleProcess}
            disabled={isProcessing}
          >
            {isProcessing ? (
              <>
                <Loader2 className="w-5 h-5 animate-spin" />
                Processing Files...
              </>
            ) : (
              <>
                <FileSpreadsheet className="w-5 h-5" />
                Process {selectedFiles.length} File{selectedFiles.length > 1 ? 's' : ''}
              </>
            )}
          </Button>
        </div>
      )}
    </div>
  );
}
