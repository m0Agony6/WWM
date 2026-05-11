import React, { useState, useRef, useEffect } from 'react';
import { 
  FileSpreadsheet, 
  Upload, 
  Settings, 
  MessageSquare, 
  Play, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Loader2,
  ChevronRight,
  Shield,
  FileBox,
  X,
  Plus,
  Trash2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { AppConfig, ChatMessage } from './types';

export default function App() {
  const [step, setStep] = useState(1);
  const [ledger, setLedger] = useState<File | null>(null);
  const [templateLeft, setTemplateLeft] = useState<File | null>(null);
  const [templateRight, setTemplateRight] = useState<File | null>(null);
  const [serials, setSerials] = useState("");
  const [sourceSheet, setSourceSheet] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [resultUrl, setResultUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Config State
  const [config, setConfig] = useState<AppConfig>({
    reportPrefixes: {
      "G8": "07062500C700308",
      "G9": "07062500C700309",
      "G10": "07062500C700311",
      "G11": "07062500C700310",
      "G12": "07062500C500100",
      "G15": "07062500C511100",
    },
    supervisors: {
      "正环（建-名）（远大）+（十八局）（江南）": "中铁华铁工程设计集团有限公司石家庄市城市轨道交通5号线一期工程监理5202标段监理部",
      "正环（十八局）（江南）+（二十局）（华铁）+解-建（远大）": "浙江江南工程管理股份有限公司石家庄市城市轨道交通5号线一期工程监理5201标段监理部",
    },
    moldPrefixes: {
      "D4": "B1",
      "F4": "B2",
      "H4": "B3",
      "J4": "L1",
      "L4": "L2",
      "O4": "F",
    }
  });

  // Chat State
  const [showChat, setShowChat] = useState(false);
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState("");
  const chatEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  const handleProcess = async () => {
    if (!ledger || !templateLeft || !templateRight || !serials) return;
    
    setIsProcessing(true);
    setError(null);
    setResultUrl(null);

    const formData = new FormData();
    formData.append('ledger', ledger);
    formData.append('templateLeft', templateLeft);
    formData.append('templateRight', templateRight);
    formData.append('data', JSON.stringify({
      serials,
      config,
      sourceSheet
    }));

    try {
      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error || 'Processing failed');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setResultUrl(url);
      setStep(3);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleSendMessage = async () => {
    if (!input.trim()) return;
    const userMsg = { role: 'user' as const, text: input };
    setMessages(prev => [...prev, userMsg]);
    setInput("");

    try {
      const response = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          message: input,
          context: { ledger: ledger?.name, serials, step }
        }),
      });
      const data = await response.json();
      setMessages(prev => [...prev, { role: 'ai', text: data.reply }]);
    } catch (err) {
      setMessages(prev => [...prev, { role: 'ai', text: "抱歉，AI连接出现问题。" }]);
    }
  };

  const FileUpload = ({ file, setFile, label, icon: Icon }: any) => (
    <div className="flex flex-col gap-2">
      <label className="text-sm font-medium text-slate-700 flex items-center gap-2">
        <Icon className="w-4 h-4 text-slate-400" />
        {label}
      </label>
      <div 
        className={`relative group border-2 border-dashed rounded-xl p-6 transition-all duration-300 flex flex-col items-center justify-center gap-3
          ${file ? 'border-emerald-200 bg-emerald-50/30' : 'border-slate-200 hover:border-indigo-400 hover:bg-slate-50'}`}
      >
        <input 
          type="file" 
          accept=".xlsx,.xlsm"
          onChange={(e) => setFile(e.target.files?.[0] || null)}
          className="absolute inset-0 opacity-0 cursor-pointer"
        />
        {file ? (
          <>
            <div className="w-12 h-12 rounded-full bg-emerald-100 flex items-center justify-center">
              <CheckCircle2 className="w-6 h-6 text-emerald-600" />
            </div>
            <div className="text-center">
              <p className="text-sm font-semibold text-emerald-900 truncate max-w-[200px]">{file.name}</p>
              <p className="text-xs text-emerald-600">已就绪</p>
            </div>
          </>
        ) : (
          <>
            <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center group-hover:bg-indigo-100 transition-colors">
              <Upload className="w-6 h-6 text-slate-400 group-hover:text-indigo-600" />
            </div>
            <div className="text-center">
              <p className="text-sm font-medium text-slate-600">点击或并拖拽上传</p>
              <p className="text-xs text-slate-400">支持 .xlsx, .xlsm 格式</p>
            </div>
          </>
        )}
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans selection:bg-indigo-100">
      {/* Header */}
      <header className="sticky top-0 z-40 w-full border-b border-slate-200 bg-white/80 backdrop-blur-md">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white shadow-lg shadow-indigo-200">
              <Shield className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight text-slate-900">管片资料自动化工具</h1>
              <p className="text-[10px] uppercase tracking-wider font-bold text-indigo-600">Precision Engineering Automator</p>
            </div>
          </div>
          
          <div className="flex items-center gap-4">
            <button 
              onClick={() => setShowChat(!showChat)}
              className="p-2 hover:bg-slate-100 rounded-lg transition-colors relative"
            >
              <MessageSquare className="w-5 h-5 text-slate-600" />
              {messages.length > 0 && (
                <span className="absolute top-1 right-1 w-2 h-2 bg-indigo-600 rounded-full border-2 border-white" />
              )}
            </button>
            <div className="h-4 w-px bg-slate-200" />
            <div className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 rounded-full">
              <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse" />
              <span className="text-xs font-medium text-slate-600">系统就绪</span>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Left Column: Progress & Steps */}
        <div className="lg:col-span-8 space-y-6">
          {/* Progress Tracker */}
          <div className="flex items-center gap-6 mb-8 overflow-x-auto pb-4 scrollbar-hide">
            {[1, 2, 3].map((s) => (
              <div key={s} className="flex items-center gap-3 group">
                <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold border-2 transition-all duration-500
                  ${step === s ? 'bg-indigo-600 border-indigo-600 text-white shadow-lg shadow-indigo-100 scale-110' : 
                    step > s ? 'bg-emerald-500 border-emerald-500 text-white' : 'bg-white border-slate-200 text-slate-400'}`}
                >
                  {step > s ? <CheckCircle2 className="w-5 h-5" /> : s}
                </div>
                <span className={`text-sm font-semibold whitespace-nowrap transition-colors duration-500 ${step === s ? 'text-slate-900' : 'text-slate-400'}`}>
                  {s === 1 ? '文件准备' : s === 2 ? '参数设定' : '结果下载'}
                </span>
                {s < 3 && <ChevronRight className="w-4 h-4 text-slate-300" />}
              </div>
            ))}
          </div>

          <AnimatePresence mode="wait">
            {step === 1 && (
              <motion.div 
                key="step1"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="bg-white rounded-2xl border border-slate-200 p-8 shadow-sm"
              >
                <div className="mb-8">
                  <h2 className="text-xl font-bold text-slate-900 mb-1">资料库上传</h2>
                  <p className="text-sm text-slate-500">上传基础台账与双向管片模板</p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  <FileUpload file={ledger} setFile={setLedger} label="管片台账 (Ledger)" icon={FileSpreadsheet} />
                  <FileUpload file={templateLeft} setFile={setTemplateLeft} label="左转模板 (Template L)" icon={FileBox} />
                  <FileUpload file={templateRight} setFile={setTemplateRight} label="右转模板 (Template R)" icon={FileBox} />
                </div>

                <div className="mt-12 flex justify-end">
                  <button 
                    disabled={!ledger || !templateLeft || !templateRight}
                    onClick={() => setStep(2)}
                    className="group bg-indigo-600 disabled:bg-slate-200 text-white px-6 py-3 rounded-xl font-bold transition-all duration-300 flex items-center gap-2 hover:bg-indigo-700 hover:shadow-xl hover:shadow-indigo-100 active:scale-95"
                  >
                    下一步
                    <ChevronRight className="w-4 h-4 group-hover:translate-x-1 transition-transform" />
                  </button>
                </div>
              </motion.div>
            )}

            {step === 2 && (
              <motion.div 
                key="step2"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="bg-white rounded-2xl border border-slate-200 p-8 shadow-sm"
              >
                <div className="mb-8">
                  <h2 className="text-xl font-bold text-slate-900 mb-1">转换参数设定</h2>
                  <p className="text-sm text-slate-500">指定批量范围与数据路径</p>
                </div>

                <div className="space-y-6">
                  <div className="space-y-2">
                    <label className="text-sm font-bold text-slate-700">资料流水号</label>
                    <textarea 
                      placeholder="例如: 2575, 5017 或 101-150"
                      value={serials}
                      onChange={(e) => setSerials(e.target.value)}
                      className="w-full h-32 px-4 py-3 rounded-xl border border-slate-200 focus:border-indigo-500 focus:ring-4 focus:ring-indigo-50 focus:outline-none transition-all resize-none text-slate-600 font-mono text-sm"
                    />
                    <p className="text-xs text-slate-400">支持逗号、空格、回车分隔。系统将自动匹配台账记录。</p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2">
                      <label className="text-sm font-bold text-slate-700">指定工作表 (可选)</label>
                      <input 
                        type="text"
                        placeholder="留空即搜索全部"
                        value={sourceSheet}
                        onChange={(e) => setSourceSheet(e.target.value)}
                        className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:border-indigo-500 focus:ring-4 focus:ring-indigo-50 focus:outline-none transition-all"
                      />
                    </div>
                  </div>
                </div>

                {error && (
                  <div className="mt-6 p-4 bg-rose-50 border border-rose-100 rounded-xl flex items-start gap-3">
                    <AlertCircle className="w-5 h-5 text-rose-500 mt-0.5" />
                    <p className="text-sm text-rose-600 leading-relaxed font-medium">{error}</p>
                  </div>
                )}

                <div className="mt-12 flex justify-between">
                  <button 
                    onClick={() => setStep(1)}
                    className="px-6 py-3 text-slate-500 font-bold hover:text-slate-900 transition-colors"
                  >
                    返回重选
                  </button>
                  <button 
                    disabled={isProcessing || !serials}
                    onClick={handleProcess}
                    className="bg-indigo-600 disabled:bg-slate-200 text-white px-8 py-3 rounded-xl font-bold shadow-lg shadow-indigo-100 transition-all active:scale-95 flex items-center gap-2 hover:bg-indigo-700"
                  >
                    {isProcessing ? (
                      <>
                        <Loader2 className="w-5 h-5 animate-spin" />
                        正在处理中...
                      </>
                    ) : (
                      <>
                        <Play className="w-4 h-4 fill-current" />
                        启动批量生成
                      </>
                    )}
                  </button>
                </div>
              </motion.div>
            )}

            {step === 3 && (
              <motion.div 
                key="step3"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="bg-white rounded-2xl border border-slate-200 p-12 text-center shadow-sm"
              >
                <div className="w-20 h-20 bg-emerald-100 rounded-3xl mx-auto flex items-center justify-center mb-6">
                  <CheckCircle2 className="w-10 h-10 text-emerald-600" />
                </div>
                <h2 className="text-2xl font-bold text-slate-900 mb-2">生成任务已完成</h2>
                <p className="text-slate-500 mb-8 max-w-md mx-auto">
                  已成功处理所有选定的流水号。由于浏览器环境无法直接写入本地系统，请点击下方按钮下载 ZIP 压缩包。
                </p>
                
                <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                  <a 
                    href={resultUrl || '#'} 
                    download="管片目标资料.zip"
                    className="w-full sm:w-auto bg-indigo-600 text-white px-8 py-4 rounded-xl font-bold flex items-center justify-center gap-3 shadow-xl shadow-indigo-100 hover:bg-indigo-700 transition-all active:scale-95"
                  >
                    <Download className="w-5 h-5" />
                    立即下载压缩包
                  </a>
                  <button 
                    onClick={() => {
                      setStep(1);
                      setResultUrl(null);
                    }}
                    className="w-full sm:w-auto px-8 py-4 text-slate-500 font-bold hover:text-slate-900 transition-colors"
                  >
                    开启新任务
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Right Column: Settings & Sidebar */}
        <div className="lg:col-span-4 space-y-6">
          <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
            <div className="p-5 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <div className="flex items-center gap-2">
                <Settings className="w-4 h-4 text-slate-500" />
                <h3 className="font-bold text-slate-900">个性化设定</h3>
              </div>
            </div>
            <div className="p-5 space-y-6 max-h-[60vh] overflow-y-auto">
              {/* Report Prefixes */}
              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">报验编号前缀 (Cell Prefix)</span>
                </div>
                <div className="space-y-2">
                  {Object.entries(config.reportPrefixes).map(([cell, value]) => (
                    <div key={cell} className="flex gap-2 group">
                      <div className="w-12 px-2 py-1.5 bg-slate-100 rounded-lg text-[10px] font-mono font-bold text-slate-500 flex items-center justify-center">
                        {cell}
                      </div>
                      <input 
                        type="text"
                        value={value}
                        onChange={(e) => {
                          const newPrefixes = { ...config.reportPrefixes, [cell]: e.target.value };
                          setConfig({ ...config, reportPrefixes: newPrefixes });
                        }}
                        className="flex-1 px-3 py-1.5 text-xs rounded-lg border border-slate-200 focus:outline-none focus:border-indigo-400 font-mono text-slate-600"
                      />
                    </div>
                  ))}
                </div>
              </div>

              {/* Mold Prefixes */}
              <div className="space-y-4 pt-4 border-t border-slate-100">
                <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">模具编号配置</span>
                <div className="space-y-2">
                  {Object.entries(config.moldPrefixes).map(([cell, value]) => (
                    <div key={cell} className="flex gap-2">
                      <div className="w-12 px-2 py-1.5 bg-slate-100 rounded-lg text-[10px] font-mono font-bold text-slate-500 flex items-center justify-center">
                        {cell}
                      </div>
                      <input 
                        type="text"
                        value={value}
                        onChange={(e) => {
                          const newMold = { ...config.moldPrefixes, [cell]: e.target.value };
                          setConfig({ ...config, moldPrefixes: newMold });
                        }}
                        className="flex-1 px-3 py-1.5 text-xs rounded-lg border border-slate-200 focus:outline-none focus:border-indigo-400 font-mono text-slate-600"
                      />
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>

          <div className="p-6 bg-indigo-900 rounded-2xl text-white relative overflow-hidden group">
            <div className="relative z-10">
              <h3 className="font-bold mb-1">正在处理复杂数据？</h3>
              <p className="text-sm text-indigo-200 mb-4 leading-relaxed">
                点击右下角对话图标，让 AI 助手帮您校验流水号或分析台账结构。
              </p>
              <button 
                onClick={() => setShowChat(true)}
                className="text-xs font-bold bg-white/10 hover:bg-white/20 px-3 py-2 rounded-lg transition-colors flex items-center gap-2"
              >
                立即咨询助理
                <ChevronRight className="w-3 h-3" />
              </button>
            </div>
            <div className="absolute -right-4 -bottom-4 w-32 h-32 bg-indigo-500/10 rounded-full blur-2xl group-hover:scale-150 transition-transform duration-700" />
          </div>
        </div>
      </main>

      {/* Floating Chat */}
      <AnimatePresence>
        {showChat && (
          <motion.div 
            initial={{ opacity: 0, y: 100, scale: 0.9 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 100, scale: 0.9 }}
            className="fixed bottom-6 right-6 w-96 max-w-[calc(100vw-3rem)] h-[560px] bg-white rounded-3xl shadow-2xl shadow-indigo-200 border border-slate-200 z-50 flex flex-col overflow-hidden"
          >
            <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-indigo-600 text-white">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-white/20 rounded-lg">
                  <MessageSquare className="w-4 h-4" />
                </div>
                <span className="font-bold text-sm">资料辅助 AI 助手</span>
              </div>
              <button onClick={() => setShowChat(false)} className="p-1 hover:bg-white/20 rounded-md transition-colors">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="flex-1 overflow-y-auto p-4 space-y-4">
              {messages.length === 0 && (
                <div className="h-full flex flex-col items-center justify-center text-center p-8">
                  <div className="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mb-4">
                    <MessageSquare className="w-8 h-8 text-slate-300" />
                  </div>
                  <p className="text-sm font-bold text-slate-800 mb-1">您好，我是资料助手</p>
                  <p className="text-xs text-slate-500">您可以问我如何映射表格，或者让我帮您检查流水号范围。</p>
                </div>
              )}
              {messages.map((m, i) => (
                <div key={i} className={`flex ${m.role === 'user' ? 'justify-end' : 'justify-start'}`}>
                  <div className={`max-w-[85%] rounded-2xl px-4 py-2.5 text-sm leading-relaxed ${
                    m.role === 'user' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'
                  }`}>
                    {m.text}
                  </div>
                </div>
              ))}
              <div ref={chatEndRef} />
            </div>

            <div className="p-4 border-t border-slate-100 bg-slate-50">
              <div className="flex gap-2">
                <input 
                  type="text"
                  placeholder="请输入您的问题..."
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  onKeyPress={(e) => e.key === 'Enter' && handleSendMessage()}
                  className="flex-1 px-4 py-2 rounded-xl border border-slate-200 focus:outline-none focus:border-indigo-400 text-sm"
                />
                <button 
                  onClick={handleSendMessage}
                  disabled={!input.trim()}
                  className="p-2 bg-indigo-600 text-white rounded-xl disabled:bg-slate-300 transition-colors"
                >
                  <ChevronRight className="w-5 h-5" />
                </button>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
