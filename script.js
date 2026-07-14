// ==========================================
// 1. DATA MATRIX & COUNTRY MAP (From Excel)
// ==========================================
const RATES = {"0.5": [319000, 343000, 369000, 434000, 574000, 525000, 608000, 820000], "1": [407000, 460000, 494000, 574000, 728000, 728000, 826000, 1101000], "1.5": [462000, 545000, 579000, 670000, 856000, 855000, 964000, 1274000], "2": [517000, 630000, 664000, 766000, 984000, 982000, 1103000, 1448000], "2.5": [574000, 715000, 749000, 862000, 1112000, 1109000, 1234000, 1610000], "3": [626000, 786000, 828000, 950000, 1221000, 1221000, 1351000, 1758000], "3.5": [679000, 858000, 906000, 1039000, 1331000, 1333000, 1467000, 1906000], "4": [731000, 930000, 985000, 1127000, 1440000, 1445000, 1584000, 2054000], "4.5": [784000, 1002000, 1064000, 1216000, 1550000, 1558000, 1700000, 2202000], "5": [836000, 1074000, 1143000, 1304000, 1659000, 1670000, 1817000, 2350000], "5.5": [885000, 1146000, 1222000, 1393000, 1778000, 1782000, 1933000, 2499000], "6": [935000, 1218000, 1301000, 1481000, 1897000, 1894000, 2050000, 2647000], "6.5": [984000, 1289000, 1380000, 1570000, 2017000, 2006000, 2166000, 2795000], "7": [1033000, 1361000, 1458000, 1658000, 2136000, 2118000, 2283000, 2943000], "7.5": [1082000, 1433000, 1537000, 1747000, 2255000, 2230000, 2400000, 3091000], "8": [1131000, 1505000, 1616000, 1835000, 2374000, 2343000, 2516000, 3239000], "8.5": [1180000, 1577000, 1695000, 1924000, 2493000, 2455000, 2633000, 3387000], "9": [1229000, 1649000, 1774000, 2012000, 2612000, 2567000, 2749000, 3535000], "9.5": [1278000, 1720000, 1853000, 2101000, 2732000, 2679000, 2866000, 3683000], "10": [1327000, 1792000, 1932000, 2189000, 2851000, 2791000, 2982000, 3831000], "10.5": [1374000, 1859000, 2002000, 2272000, 2974000, 2903000, 3099000, 3979000], "11": [1420000, 1926000, 2072000, 2354000, 3098000, 3016000, 3215000, 4127000], "11.5": [1466000, 1992000, 2142000, 2436000, 3221000, 3128000, 3332000, 4275000], "12": [1513000, 2059000, 2212000, 2519000, 3345000, 3240000, 3448000, 4424000], "12.5": [1559000, 2125000, 2282000, 2601000, 3468000, 3352000, 3565000, 4572000], "13": [1606000, 2192000, 2352000, 2683000, 3592000, 3464000, 3681000, 4720000], "13.5": [1652000, 2258000, 2422000, 2766000, 3716000, 3576000, 3798000, 4868000], "14": [1699000, 2325000, 2492000, 2848000, 3839000, 3688000, 3914000, 5160000], "14.5": [1745000, 2392000, 2563000, 2931000, 3963000, 3801000, 4031000, 5164000], "15": [1791000, 2458000, 2633000, 3013000, 4086000, 3913000, 4148000, 5312000], "15.5": [1838000, 2525000, 2703000, 3095000, 4210000, 4025000, 4264000, 5460000], "16": [1884000, 2591000, 2773000, 3178000, 4333000, 4137000, 4381000, 5608000], "16.5": [1931000, 2658000, 2843000, 3260000, 4457000, 4249000, 4497000, 5756000], "17": [1977000, 2725000, 2913000, 3342000, 4580000, 4361000, 4614000, 5904000], "17.5": [2024000, 2791000, 2983000, 3425000, 4704000, 4474000, 4730000, 6052000], "18": [2070000, 2858000, 3053000, 3507000, 4827000, 4586000, 4847000, 6200000], "18.5": [2117000, 2924000, 3123000, 3589000, 4951000, 4698000, 4963000, 6349000], "19": [2163000, 2991000, 3193000, 3672000, 5075000, 4810000, 5080000, 6497000], "19.5": [2209000, 3058000, 3263000, 3754000, 5198000, 4922000, 5196000, 6645000], "20": [2256000, 3124000, 3334000, 3837000, 5322000, 5034000, 5313000, 6793000], "20.5": [2295000, 3180000, 3394000, 3907000, 5436000, 5122000, 5411000, 6918000], "21": [2335000, 3236000, 3454000, 3978000, 5549000, 5210000, 5509000, 7043000], "21.5": [2374000, 3292000, 3515000, 4049000, 5663000, 5297000, 5607000, 7169000], "22": [2414000, 3348000, 3575000, 4120000, 5777000, 5385000, 5705000, 7294000], "22.5": [2453000, 3405000, 3636000, 4191000, 5891000, 5472000, 5804000, 7419000], "23": [2492000, 3461000, 3696000, 4262000, 6005000, 5560000, 5902000, 7545000], "23.5": [2532000, 3517000, 3757000, 4333000, 6119000, 5648000, 6000000, 7670000], "24": [2571000, 3573000, 3817000, 4404000, 6233000, 5735000, 6098000, 7795000], "24.5": [2611000, 3629000, 3878000, 4475000, 6347000, 5823000, 6196000, 7920000], "25": [2650000, 3685000, 3938000, 4546000, 6461000, 5910000, 6294000, 8046000], "25.5": [2690000, 3741000, 3999000, 4617000, 6575000, 5998000, 6392000, 8171000], "26": [2729000, 3797000, 4059000, 4688000, 6689000, 6086000, 6491000, 8296000], "26.5": [2768000, 3853000, 4120000, 4759000, 6802000, 6173000, 6589000, 8422000], "27": [2808000, 3909000, 4180000, 4830000, 6916000, 6261000, 6687000, 8547000], "27.5": [2847000, 3965000, 4240000, 4901000, 7030000, 6349000, 6785000, 8672000], "28": [2887000, 4021000, 4301000, 4972000, 7144000, 6436000, 6883000, 8798000], "28.5": [2926000, 4077000, 4361000, 5043000, 7258000, 6524000, 6981000, 8923000], "29": [2966000, 4134000, 4422000, 5114000, 7372000, 6611000, 7079000, 9048000], "29.5": [3005000, 4190000, 4482000, 5185000, 7486000, 6699000, 7177000, 9173000], "30": [3044000, 4246000, 4543000, 5256000, 7600000, 6787000, 7276000, 9299000]};

const COUNTRIES = {"Afghanistan (AF)": 8, "Albania (AL)": 8, "Algeria (DZ)": 8, "Amenia (AM)": 8, "American Samoa (AS)": 8, "Andorra (AD)": 7, "Angola (AO)": 8, "Anguilla (AI)": 8, "Antigua (AG)": 8, "Argentina (AR)": 8, "Aruba (AW)": 8, "Australia (AU)": 4, "Austria (AT)": 7, "Azerbaijan (AZ)": 8, "Bahamas (BS)": 8, "Bahrain (BH)": 6, "Bangladesh (BD)": 6, "Barbados (BB)": 8, "Belarus (BY)": 8, "Belgium (BE)": 7, "Belize (BZ)": 8, "Benin (BJ)": 8, "Bermuda (BM)": 8, "Bhutan (BT)": 8, "Bolivia (BO)": 8, "Bonaire (XB)": 8, "Bosnia & Harzegovina (BA)": 8, "Botswana (BW)": 8, "Brazil (BR)": 8, "Brunei (BN)": 2, "Bulgaria (BG)": 7, "Burkina Faso (BF)": 8, "Burundi (BI)": 8, "Cambodia (KH)": 2, "Cameroon (CM)": 8, "Canada (CA)": 5, "Canary Islands,The (IC)": 8, "Cape Verde (CV)": 8, "Cayman Islands (KY)": 8, "Central African Rep (CF)": 8, "Chad (TD)": 8, "Chile (CL)": 8, "China (CN) - Rest of China": 4, "China (CN) - South China (Guangzhou, Shenzhen, etc.)": 3, "Cole D Invoice (CI)": 8, "Colombia (CO)": 8, "Comoros (KM)": 8, "Congo (CG)": 8, "Congo,DPR (CD)": 8, "Cook Islands (CK)": 8, "Costa Rica (CR)": 8, "Croatia (HR)": 7, "Cuba (CU)": 8, "Curacao (XC)": 8, "Cyprus (CY)": 7, "Czech Rep, The (GZ)": 7, "Denmark (DK)": 7, "Djibouti (DJ)": 8, "Dominian Rep. (DO)": 8, "Ecuador (EC)": 8, "Egypt (EG)": 8, "El Salvador (SV)": 8, "Eritrea (ER)": 8, "Estonia (EE)": 7, "Eswatini (SZ)": 8, "Ethiopia (ET)": 8, "Faikland Islands (FK)": 8, "Faroe Islands (FO)": 8, "Fiji (FI)": 8, "Finland (FI)": 7, "France (FR)": 7, "French Guyana (GF)": 8, "Gabon (GA)": 8, "Gambia (GM)": 8, "Georgia (GE)": 8, "Germany (DE)": 7, "Ghana (GH)": 8, "Gibraltar (GI)": 8, "Greece (GR)": 7, "Greenland (GL)": 8, "Grenada (GD)": 8, "Guadelope (GP)": 8, "Guam (GU)": 8, "Guatemala (GU)": 8, "Guernsey (GG)": 8, "Guinea - Bissau (GW)": 8, "Guinea - Equatorial (GQ)": 8, "Guinea - Rep. (GN)": 8, "Guyana (British) (GY)": 8, "Haiti (HT)": 8, "Honduras (HN)": 8, "Hong Kong SAR China (HK)": 2, "Hungary (HU)": 7, "Iceland (IS)": 8, "India (in)": 6, "Iran (IR)": 8, "Iraq (IQ)": 8, "Ireland, Rep. Of (IE)": 7, "Israel (IL)": 8, "Italy (IT)": 7, "Jamaica (JM)": 8, "Japan (JP)": 3, "Jersey (JE)": 8, "Jordan (JO)": 6, "Kazakkhstan (KZ)": 8, "Kenya (KE)": 8, "Kiribati (KI)": 8, "Korea, D.P.R Of (KP)": 8, "Korea, Rep, Of (KR)": 4, "Kosovo (KV)": 8, "Kuwait (KW)": 6, "Kyrgyzstan (KG)": 8, "Lalvia (LV)": 7, "Laos (LA)": 2, "Lebanon (LB)": 8, "Lesotho (LS)": 8, "Libya (LY)": 8, "Liechtenstein (LI)": 7, "Lithuania (LT)": 7, "Luxembourg (LU)": 7, "Macau SAR China (MO)": 2, "Madagascar (MG)": 8, "Malawi (MW)": 8, "Malaysia (MY)": 2, "Maldives (MV)": 6, "Mali (ML)": 8, "Malta (MT)": 7, "Marshall Islands (MH)": 8, "Martinique (MQ)": 8, "Mauritania (MR)": 8, "Mauritius (MU)": 8, "Mayotte (YT)": 8, "Mexico (MX)": 5, "Micronesia (FM)": 8, "Moldova, Rep, Of (MD)": 8, "Monaco (MC)": 7, "Mongalia (MN)": 8, "Monstserrat (MS)": 8, "Montenegro, Rep Of (ME)": 8, "Morocco (MA)": 8, "Mozambique (MZ)": 8, "Myanmar (MM)": 2, "Namibia (NA)": 8, "Nauru, Rep, Of (NR)": 8, "Nepal (NP)": 6, "Netherlands, The (NL)": 7, "Nevis (XN)": 8, "New Caledonia (NC)": 8, "New Zealand (NZ)": 4, "Nicaragua (NI)": 8, "Niger (NE)": 8, "Nigeria (NG)": 8, "Niue (NU)": 8, "North Macedonia (MK)": 8, "Northern Mariana Islands (MP)": 8, "Norway (NO)": 7, "Oman (OM)": 6, "Pakistan (PK)": 6, "Palau (PW)": 8, "Panama (PA)": 8, "Papua New Guinea (PG)": 4, "Paraguay (PY)": 8, "Peru (PE)": 8, "Philippines, The (PH)": 2, "Poland (PL)": 7, "Portugal (PT)": 7, "Puerto Rico (PR)": 8, "Qatar (QA)": 6, "Reunion, Island Of (RE)": 8, "Romania (RO)": 7, "Russian Federation (RU)": 8, "Rwanda (RW)": 8, "Saint Helena (SH)": 8, "Samoa (WS)": 8, "San Marino  (SM)": 7, "Sao Tome and Principe  (ST)": 8, "Saudi Arabia (SA)": 6, "Senegal (SN)": 8, "Serbia, Rep. Of (RS)": 8, "Seychelles (SC)": 8, "Sierra Leone (SL)": 8, "Singapore (SG)": 1, "Slovakia (SK)": 7, "Slovenia (SI)": 7, "Solomon Islands (SB)": 8, "Somalia (SO)": 8, "Somaliland, Rep Of (XS)": 8, "South Africa (ZA)": 7, "South Saudan (SS)": 8, "Spain (ES)": 7, "Sri Lanka (LK)": 6, "St. Barthelemy (XY)": 8, "St. Eustatius (XE)": 8, "St. Kitts (KN)": 8, "St. Lucia (LC)": 8, "St. Maarten (XM)": 8, "St. Vincent (VC)": 8, "Sudan (SD)": 8, "Suriname (SR)": 8, "Sweden (SE)": 7, "Switzerland (CH)": 7, "Syria (SY)": 8, "Tahiti (PF)": 8, "Taiwan (TW)": 4, "Tajikistan (TJ)": 8, "Tanzania (TZ)": 8, "Thailand (TH)": 2, "Timor-Leste (TL)": 1, "Togo (TG)": 8, "Tonga (TO)": 8, "Trinidad And Tabago (TT)": 8, "Tunisia (TN)": 8, "Turkey (TR)": 7, "Turkmenistan (TM)": 8, "Turks & Caicos (TC)": 8, "Tuvalu (TV)": 8, "USA (US)": 5, "Uganda (UG)": 8, "Ukraine (UA)": 8, "United Arab Emitayrs (AE)": 6, "United Kingdom (GB)": 7, "Uruguay (UY)": 8, "Uzbekistan (UZ)": 8, "Vanuatu (VU)": 8, "Vatica City (VA)": 7, "Venezuela (VE)": 8, "Vietnam (VN)": 2, "Virgin Island - US (VI)": 8, "Virgin Islands-British (VG)": 8, "Yemen, Rep. Of (YE)": 8, "Zambia (ZM)": 8, "Zimbabwe  (ZW)": 8};

// ==========================================
// 2. DOM ELEMENT SELECTORS
// ==========================================
const toDestination = document.getElementById('to');
const routeChip = document.getElementById('routeChip');
const estimateBtn = document.getElementById('estimateBtn');
const resetBtn = document.getElementById('reset');
const remoteAreaCheck = document.getElementById('remoteAreaCheck');

// ==========================================
// 3. UI INITIALIZATION & SYNC
// ==========================================
function populateCountries() {
  toDestination.innerHTML = '';
  Object.keys(COUNTRIES).forEach(country => {
    const opt = document.createElement('option');
    opt.value = country;
    opt.textContent = country;
    toDestination.appendChild(opt);
  });
}

function formatIDR(amount) {
  return 'Rp ' + Math.round(amount).toLocaleString('id-ID');
}

function updateRouteChip() {
  const fromText = document.getElementById('from').selectedOptions[0].textContent.split(',')[0];
  const toText = toDestination.value || "Destination";
  routeChip.textContent = `${fromText} → ${toText}`;
  
  // Update live preview indicator block value
  const zoneNum = COUNTRIES[toDestination.value] || 8;
  document.getElementById('zoneClassificationVal').textContent = `Zone ${zoneNum}`;
}

// ==========================================
// 4. RATE ENGINE CALCULATIONS
// ==========================================
function calculateEstimate() {
  const targetCountry = toDestination.value;
  if (!targetCountry) return;

  const zoneIdx = COUNTRIES[targetCountry] - 1; // Maps Zone 1-8 array column indices
  const inputWeight = parseFloat(document.getElementById('weight').value) || 0.5;
  const lengthCm = parseFloat(document.getElementById('length').value) || 0;
  const widthCm = parseFloat(document.getElementById('width').value) || 0;
  const heightCm = parseFloat(document.getElementById('height').value) || 0;

  // Process Volumetric check bounds
  let volumetricWeight = 0;
  if (lengthCm > 0 && widthCm > 0 && heightCm > 0) {
    volumetricWeight = (lengthCm * widthCm * heightCm) / 5000;
  }

  // Determine target weight tier factor (Round steps up to nearest 0.5 below 30kg threshold)
  let finalWeight = Math.max(inputWeight, volumetricWeight);
  if (finalWeight <= 30) {
    finalWeight = Math.ceil(finalWeight * 2) / 2;
  } else {
    finalWeight = Math.ceil(finalWeight);
  }

  document.getElementById('chargeableWeightVal').textContent = `${finalWeight} kg ${volumetricWeight > inputWeight ? '(Volumetric Weight Applied)' : ''}`;

  let freight = 0;
  let surcharge = 0;

  if (finalWeight > 30) {
    alert(`Cargo weight status (${finalWeight} kg) exceeds primary 30 KG baseline grid tier. Calculating based on max 30 KG row baseline factor.`);
    const maxTierRow = RATES["30"];
    freight = finalWeight * (maxTierRow[zoneIdx] / 30);
  } else {
    const searchKey = finalWeight.toString();
    if (RATES[searchKey]) {
      freight = RATES[searchKey][zoneIdx];
    } else {
      freight = RATES["30"][zoneIdx];
    }
  }

  // Handle service tier level shifts (+40% match condition mapping to sheet calculations)
  if (document.getElementById('service').value === 'express') {
    freight *= 1.40;
  }

  // Remote Handling Surcharges check addition logic
  if (remoteAreaCheck.checked) {
    surcharge += 450000;
  }

  // Display outputs formatted variables injection
  document.getElementById('freightVal').textContent = formatIDR(freight);
  document.getElementById('surchargeVal').textContent = formatIDR(surcharge);
  document.getElementById('totalVal').textContent = formatIDR(freight + surcharge);
}

// ==========================================
// 5. ATTACH HANDLERS & LIFECYCLE
// ==========================================
toDestination.addEventListener('change', updateRouteChip);
document.getElementById('from').addEventListener('change', updateRouteChip);
estimateBtn.addEventListener('click', calculateEstimate);

resetBtn.addEventListener('click', () => {
  document.getElementById('weight').value = '5';
  document.getElementById('length').value = '';
  document.getElementById('width').value = '';
  document.getElementById('height').value = '';
  document.getElementById('service').value = 'standard';
  toDestination.selectedIndex = 0;
  remoteAreaCheck.checked = false;
  document.getElementById('chargeableWeightVal').textContent = '5 kg';
  ['freightVal', 'surchargeVal', 'totalVal'].forEach(id => document.getElementById(id).textContent = '-');
  updateRouteChip();
});

// Run setup init procedures
populateCountries();
updateRouteChip();