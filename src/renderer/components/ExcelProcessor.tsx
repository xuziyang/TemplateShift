import { useState } from 'react';
import { Box, Button, Typography } from '@mui/material';
import * as XLSX from 'xlsx';

interface ExcelData {
  companyName: string; // 公司名称
  registrationStatus: string; // 登记状态
  legalRepresentative: string; // 法定代表人
  province: string; // 所属省份
  city: string; // 所属城市
  validPhone: string; // 有效手机号
  morePhones: string; // 更多电话
  registeredAddress: string; // 注册地址
  [key: string]: string;
}

export default function ExcelProcessor() {
  const [isDragging, setIsDragging] = useState(false);
  const [data, setData] = useState<ExcelData[]>([]);

  const processExcelFile = (file: File) => {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
      alert('请上传Excel文件(.xlsx或.xls格式)');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // 获取表头行并创建映射关系
        const rawHeaderRow = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[1]; // 修改为索引1，获取第二行作为表头
        if (!Array.isArray(rawHeaderRow)) {
          throw new Error('无法读取表头行');
        }
        const headerRow = rawHeaderRow as string[];
        const headerMapping: { [key: string]: string } = {
          '公司名称': 'companyName',
          '登记状态': 'registrationStatus',
          '法定代表人': 'legalRepresentative',
          '所属省份': 'province',
          '所属城市': 'city',
          '有效手机号': 'validPhone',
          '更多电话': 'morePhones',
          '注册地址': 'registeredAddress'
        };
        
        // 根据表头创建字段映射
        const headers = headerRow.map((header: string) => headerMapping[header] || header);
        const rawData = XLSX.utils.sheet_to_json<ExcelData>(worksheet, { 
          header: headers,
          range: 2, // 从第三行开始读取数据（索引为2）
          defval: '' // 设置空单元格的默认值
        });
        
        // 处理数据，优先使用有效手机号，其次是phone字段，最后是更多电话中的第一个号码
        const processedData = rawData.map(row => {
          if (!row || typeof row !== 'object') return null;
          
          // 获取最终使用的手机号，优先使用有效手机号，其次是更多电话中的号码
          const phones = [
            row.validPhone,
            ...(row.morePhones ? row.morePhones.split(',').map(p => p.trim()) : [])
          ].filter(Boolean);

          
          // 区分手机号和座机号
          const mobilePhone = phones.find(p => /^1[3-9]\d{9}$/.test(p));
          const landlinePhone = phones.find(p => /^\d{3,4}-?\d{7,8}$/.test(p));
          const finalPhone = (mobilePhone || landlinePhone || '').trim();
          
          // 处理地区信息
          const province = (row.province || '').trim();
          const city = (row.city || '').trim();
          const region = [province, city].filter(Boolean).join('-');
          
          // 使用法定代表人作为客户名
          const customerName = (row.legalRepresentative || '').trim();
          const companyName = (row.companyName || '').trim();
          
          return {
            customerName,
            companyName: companyName,
            registrationStatus: (row.registrationStatus || '').trim(),
            phone: finalPhone,
            region,
            validPhone: (row.validPhone || '').trim(),
            morePhones: (row.morePhones || '').trim(),
            province,
            city,
            registeredAddress: (row.registeredAddress || '').trim()
          };
        }).filter(row => row !== null && row.phone !== '');
        
        // 使用Map进行电话号码去重，保留最完整的记录
        const phoneMap = new Map<string, ExcelData>();
        processedData.forEach(row => {
          if (!row || !row.phone) return;
          
          const existingRow = phoneMap.get(row.phone);
          if (!existingRow) {
            phoneMap.set(row.phone, {
              companyName: row.companyName,
              registrationStatus: row.registrationStatus,
              legalRepresentative: row.customerName,
              province: row.province,
              city: row.city,
              validPhone: row.validPhone,
              morePhones: row.morePhones,
              registeredAddress: row.registeredAddress,
              phone: row.phone,
              customerName: row.customerName,
              region: [row.province, row.city].filter(Boolean).join('-')
            });
          } else {
            // 合并记录，保留非空字段
            const mergedRow = {
              companyName: row.companyName || existingRow.companyName,
              registrationStatus: row.registrationStatus || existingRow.registrationStatus,
              legalRepresentative: row.customerName || existingRow.legalRepresentative,
              province: row.province || existingRow.province,
              city: row.city || existingRow.city,
              validPhone: row.validPhone || existingRow.validPhone,
              morePhones: row.morePhones || existingRow.morePhones,
              registeredAddress: row.registeredAddress || existingRow.registeredAddress,
              phone: row.phone || existingRow.phone,
              customerName: row.customerName || existingRow.legalRepresentative,
              region: [row.province || existingRow.province, row.city || existingRow.city].filter(Boolean).join('-')
            };
            phoneMap.set(row.phone, mergedRow);
          }
        });
        
        const uniqueData = Array.from(phoneMap.values());
        // 按地区排序
        const sortedData = uniqueData.sort((a, b) => a.region.localeCompare(b.region));
        setData(sortedData);

        // 自动导出转换后的文件
        const exportWorkbook = XLSX.utils.book_new();
        const exportWorksheet = XLSX.utils.aoa_to_sheet([
          // 空行（1-12行）
          ...Array(12).fill([]),
          // 表头（第13行）
          ['客户名', '手机(必填)', '微信', '来源', '地区（市级）', '备注'],
          // 数据（从第14行开始）
          ...sortedData.map(row => [
            row.legalRepresentative || row.customerName, // 优先使用法定代表人作为客户名
            row.phone,
            '', // 微信
            row.companyName, // 来源使用公司名称
            `${row.province}-${row.city}`, // 使用省份和城市信息
            ''  // 备注
          ])
        ]);
        XLSX.utils.book_append_sheet(exportWorkbook, exportWorksheet, 'Sheet1');
        // 从原始文件名生成导出文件名
        const originalFileName = file.name;
        const exportFileName = originalFileName.replace(/\.xlsx?$/, '') + '_processed.xlsx';
        XLSX.writeFile(exportWorkbook, exportFileName);
      } catch (error) {
        console.error('Error processing Excel file:', error);
        alert('Excel文件处理失败，请检查文件格式');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    processExcelFile(file);
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    setIsDragging(false);

    const file = event.dataTransfer.files[0];
    if (file) {
      processExcelFile(file);
    }
  };

  return (
    <Box sx={{ width: '100%', mt: 3 }}>
      <Box
        sx={{
          mb: 3,
          p: 4,
          border: '2px dashed',
          borderColor: isDragging ? 'primary.main' : 'grey.200',
          borderRadius: 0,
          backgroundColor: isDragging ? 'rgba(0, 122, 255, 0.04)' : 'background.paper',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          gap: 2.5,
          cursor: 'pointer',
          transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
          boxShadow: isDragging ? '0 4px 20px rgba(0, 0, 0, 0.1)' : 'none',
          transform: isDragging ? 'scale(1.02)' : 'scale(1)'
        }}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <Typography variant="h6" sx={{ 
          color: 'text.primary',
          fontWeight: 400,
          letterSpacing: '0.5px',
          fontSize: '1.125rem',
          lineHeight: 1.5,
          opacity: 0.9
        }}>
          拖放或上传文件
        </Typography>
        <Button
          variant="contained"
          component="label"
          sx={{
            minWidth: 120,
            height: 40,
            borderRadius: 0,
            boxShadow: 'none',
            '&:hover': {
              boxShadow: 'none',
              backgroundColor: 'primary.dark'
            }
          }}
        >
          选择文件
          <input type="file" hidden accept=".xlsx,.xls" onChange={handleFileUpload} />
        </Button>
      </Box>

      <Typography variant="body1" color="text.secondary" align="center" sx={{
        mt: 2,
        fontSize: '0.875rem',
        lineHeight: 1.5,
        opacity: 0.8
      }}>
        {data.length > 0 ? '文件已成功处理并导出为 processed_data.xlsx' : '请上传Excel文件'}
      </Typography>
    </Box>
  );
}