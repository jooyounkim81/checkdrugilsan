/**
 * 약제 검증 플랫폼 v2.0 - 동일주성분별 업체수 분석 모듈
 * 주성분코드 앞 4자리 기준 성분군별 업체수 및 시장집중도 분석
 */

/**
 * 동일주성분별 업체수 분석
 * @param {Array} records - 최신 약가마스터 레코드 배열
 * @returns {Object} 성분군별 업체수 분석 결과
 */
function analyzeVendorsByIngredient(records) {
  const ingredientGroups = {};
  
  // 1단계: 주성분코드 앞 4자리 기준으로 그룹핑
  records.forEach(record => {
    const code = record.주성분코드 || record.ingredient_code || '';
    if (!code || code.length < 4) return;
    
    const groupKey = code.substring(0, 4); // 앞 4자리
    
    if (!ingredientGroups[groupKey]) {
      ingredientGroups[groupKey] = {
        codes: new Set(),
        vendors: new Set(),
        names: {},
        routes: new Set(),
        singleComplex: new Set(),
        products: []
      };
    }
    
    const group = ingredientGroups[groupKey];
    
    // 코드 추가
    group.codes.add(code);
    
    // 업체명 추가 (중복 제거)
    const vendor = record.업체명 || record.company_name || '';
    if (vendor && vendor.trim()) {
      group.vendors.add(vendor.trim());
    }
    
    // 주성분명 카운트 (대표 성분명 선정용)
    const ingredientName = record.주성분명 || record.ingredient_name || '';
    if (ingredientName) {
      group.names[ingredientName] = (group.names[ingredientName] || 0) + 1;
    }
    
    // 투여경로
    const route = record.투여경로 || record.route || '';
    if (route) {
      group.routes.add(route);
    }
    
    // 단일/복합
    const singleComplex = record.단일복합 || record.single_complex || '';
    if (singleComplex) {
      group.singleComplex.add(singleComplex);
    }
    
    // 제품 정보 저장
    group.products.push({
      code: code,
      name: ingredientName,
      vendor: vendor,
      route: route,
      productName: record.제품명 || record.product_name || ''
    });
  });
  
  // 2단계: 각 성분군별 결과 생성
  const results = [];
  
  Object.keys(ingredientGroups).sort().forEach(groupKey => {
    const group = ingredientGroups[groupKey];
    
    // 대표 주성분명 선정 (가장 많이 등장하는 것)
    let representativeName = '';
    let maxCount = 0;
    Object.keys(group.names).forEach(name => {
      if (group.names[name] > maxCount) {
        maxCount = group.names[name];
        representativeName = name;
      }
    });
    
    // 업체 목록 (정렬)
    const vendorList = Array.from(group.vendors).sort();
    const vendorCount = vendorList.length;
    
    // 시장집중도 판정
    const marketConcentration = determineMarketConcentration(vendorCount);
    
    results.push({
      ingredientGroup: groupKey,
      representativeName: representativeName,
      singleComplex: Array.from(group.singleComplex).join('/'),
      totalProducts: group.codes.size,
      vendorCount: vendorCount,
      vendorList: vendorList.join(', '),
      routes: Array.from(group.routes).join(', '),
      marketConcentration: marketConcentration,
      colorCode: getMarketConcentrationColor(vendorCount),
      products: group.products
    });
  });
  
  // 업체수 많은 순으로 정렬
  results.sort((a, b) => {
    if (b.vendorCount !== a.vendorCount) {
      return b.vendorCount - a.vendorCount;
    }
    return b.totalProducts - a.totalProducts;
  });
  
  return {
    groups: results,
    summary: {
      totalGroups: results.length,
      avgVendors: results.reduce((sum, r) => sum + r.vendorCount, 0) / results.length,
      monopoly: results.filter(r => r.vendorCount === 1).length,
      oligopoly: results.filter(r => r.vendorCount >= 2 && r.vendorCount <= 4).length,
      competitive: results.filter(r => r.vendorCount >= 5 && r.vendorCount <= 9).length,
      intense: results.filter(r => r.vendorCount >= 10).length,
      maxVendorGroup: results.length > 0 ? results[0] : null
    }
  };
}

/**
 * 시장집중도 판정
 * @param {number} vendorCount - 업체수
 * @returns {string} 시장집중도 (독점/과점/경쟁/치열)
 */
function determineMarketConcentration(vendorCount) {
  if (vendorCount === 1) return '독점';
  if (vendorCount >= 2 && vendorCount <= 4) return '과점';
  if (vendorCount >= 5 && vendorCount <= 9) return '경쟁';
  if (vendorCount >= 10) return '치열';
  return '알 수 없음';
}

/**
 * 시장집중도에 따른 색상 코드 반환
 * @param {number} vendorCount - 업체수
 * @returns {string} 색상 코드 (red/orange/green/gray)
 */
function getMarketConcentrationColor(vendorCount) {
  if (vendorCount >= 10) return 'red';      // 빨강
  if (vendorCount >= 5) return 'orange';    // 주황
  if (vendorCount >= 2) return 'green';     // 연두
  return 'gray';                             // 회색
}

/**
 * 특정 코드의 업체수 및 시장집중도 조회
 * @param {string} code - 주성분코드 (9자리)
 * @param {Object} analysisResult - analyzeVendorsByIngredient 결과
 * @returns {Object} { vendorCount, marketConcentration, colorCode }
 */
function getVendorInfoForCode(code, analysisResult) {
  if (!code || code.length < 4) {
    return { vendorCount: 0, marketConcentration: '알 수 없음', colorCode: 'gray' };
  }
  
  const groupKey = code.substring(0, 4);
  const group = analysisResult.groups.find(g => g.ingredientGroup === groupKey);
  
  if (!group) {
    return { vendorCount: 0, marketConcentration: '알 수 없음', colorCode: 'gray' };
  }
  
  return {
    vendorCount: group.vendorCount,
    marketConcentration: group.marketConcentration,
    colorCode: group.colorCode
  };
}

/**
 * 동일주성분별 업체수 시트 데이터 생성
 * @param {Object} analysisResult - analyzeVendorsByIngredient 결과
 * @returns {Array} 시트 데이터 (헤더 포함)
 */
function generateVendorAnalysisSheetData(analysisResult) {
  const data = [];
  
  // 헤더
  data.push([
    '주성분군(앞4자리)',
    '대표_주성분명',
    '단일/복합',
    '총_제품수',
    '업체수',
    '업체목록',
    '투여경로',
    '시장집중도',
    '비고'
  ]);
  
  // 데이터
  analysisResult.groups.forEach(group => {
    data.push([
      group.ingredientGroup,
      group.representativeName,
      group.singleComplex,
      group.totalProducts,
      group.vendorCount,
      group.vendorList,
      group.routes,
      group.marketConcentration,
      '' // 비고
    ]);
  });
  
  return data;
}

/**
 * 업체수 기준 엑셀 셀 색상 적용
 * @param {Object} worksheet - XLSX worksheet 객체
 * @param {Object} analysisResult - analyzeVendorsByIngredient 결과
 */
function applyVendorAnalysisColors(worksheet, analysisResult) {
  if (!worksheet || !worksheet['!ref']) return;
  
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  
  // 데이터 행에 색상 적용 (헤더 제외)
  for (let R = range.s.r + 1; R <= range.e.r; R++) {
    const rowIndex = R - range.s.r - 1; // 0-based index
    if (rowIndex < 0 || rowIndex >= analysisResult.groups.length) continue;
    
    const group = analysisResult.groups[rowIndex];
    const colorCode = group.colorCode;
    
    // 색상 매핑
    let fillColor = 'FFFFFF'; // 기본 흰색
    if (colorCode === 'red') fillColor = 'FFE5E5';       // 빨강
    else if (colorCode === 'orange') fillColor = 'FFEDD5'; // 주황
    else if (colorCode === 'green') fillColor = 'DCFCE7';  // 연두
    else if (colorCode === 'gray') fillColor = 'F3F4F6';   // 회색
    
    // 행 전체에 색상 적용
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!worksheet[cellAddress]) continue;
      
      worksheet[cellAddress].s = {
        fill: {
          fgColor: { rgb: fillColor }
        }
      };
    }
  }
}

/**
 * 요약 시트에 업체수 통계 추가
 * @param {Array} summaryData - 기존 요약 시트 데이터
 * @param {Object} analysisResult - analyzeVendorsByIngredient 결과
 * @returns {Array} 업데이트된 요약 시트 데이터
 */
function addVendorStatsToSummary(summaryData, analysisResult) {
  const summary = analysisResult.summary;
  
  // 구분선
  summaryData.push(['']);
  summaryData.push(['=== 동일주성분별 업체수 분석 (v2.0) ===']);
  summaryData.push(['']);
  
  // 통계
  summaryData.push(['전체 성분군 수', summary.totalGroups]);
  summaryData.push(['평균 업체수', summary.avgVendors.toFixed(2)]);
  summaryData.push(['']);
  
  summaryData.push(['독점 성분군 수 (업체수 1개)', summary.monopoly]);
  summaryData.push(['과점 성분군 수 (업체수 2~4개)', summary.oligopoly]);
  summaryData.push(['경쟁 성분군 수 (업체수 5~9개)', summary.competitive]);
  summaryData.push(['치열 성분군 수 (업체수 10개 이상)', summary.intense]);
  summaryData.push(['']);
  
  if (summary.maxVendorGroup) {
    summaryData.push(['최대 업체수 성분군', 
      `${summary.maxVendorGroup.ingredientGroup} (${summary.maxVendorGroup.representativeName}) - ${summary.maxVendorGroup.vendorCount}개 업체`
    ]);
  }
  
  return summaryData;
}

/**
 * 행별검증 데이터에 업체수 정보 추가
 * @param {Array} validationData - 기존 행별검증 데이터
 * @param {Object} analysisResult - analyzeVendorsByIngredient 결과
 * @returns {Array} 업데이트된 행별검증 데이터
 */
function addVendorInfoToValidation(validationData, analysisResult) {
  if (!validationData || validationData.length === 0) return validationData;
  
  // 헤더에 새 컬럼 추가
  const headerRow = validationData[0];
  const codeColumnIndex = headerRow.findIndex(h => 
    h === '입력 주성분코드' || h === '주성분코드' || h.includes('코드')
  );
  
  if (codeColumnIndex === -1) return validationData;
  
  // 헤더에 '동일주성분_업체수'와 '시장집중도' 추가
  headerRow.push('동일주성분_업체수');
  headerRow.push('시장집중도');
  
  // 데이터 행에 업체수 정보 추가
  for (let i = 1; i < validationData.length; i++) {
    const row = validationData[i];
    const code = row[codeColumnIndex];
    
    const vendorInfo = getVendorInfoForCode(code, analysisResult);
    
    row.push(vendorInfo.vendorCount);
    row.push(vendorInfo.marketConcentration);
  }
  
  return validationData;
}

/**
 * 독점/과점 성분군 추출
 * @param {Object} analysisResult - analyzeVendorsByIngredient 결과
 * @returns {Array} 독점/과점 성분군 데이터 (시트용)
 */
function extractMonopolyOligopoly(analysisResult) {
  const filtered = analysisResult.groups.filter(g => g.vendorCount <= 4);
  
  const data = [];
  
  // 헤더
  data.push([
    '주성분군(앞4자리)',
    '대표_주성분명',
    '업체수',
    '업체목록',
    '총_제품수',
    '시장집중도',
    '투여경로',
    '검토 의견'
  ]);
  
  // 데이터
  filtered.forEach(group => {
    data.push([
      group.ingredientGroup,
      group.representativeName,
      group.vendorCount,
      group.vendorList,
      group.totalProducts,
      group.marketConcentration,
      group.routes,
      group.vendorCount === 1 ? '대체 불가능, 장기 계약 고려' : '제한적 선택지, 전략적 관리 필요'
    ]);
  });
  
  return data;
}

/**
 * 전역 객체로 내보내기 (브라우저 환경)
 */
if (typeof window !== 'undefined') {
  window.MarketAnalysis = {
    analyzeVendorsByIngredient,
    determineMarketConcentration,
    getMarketConcentrationColor,
    getVendorInfoForCode,
    generateVendorAnalysisSheetData,
    applyVendorAnalysisColors,
    addVendorStatsToSummary,
    addVendorInfoToValidation,
    extractMonopolyOligopoly
  };
}

// Node.js 환경 지원
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    analyzeVendorsByIngredient,
    determineMarketConcentration,
    getMarketConcentrationColor,
    getVendorInfoForCode,
    generateVendorAnalysisSheetData,
    applyVendorAnalysisColors,
    addVendorStatsToSummary,
    addVendorInfoToValidation,
    extractMonopolyOligopoly
  };
}
