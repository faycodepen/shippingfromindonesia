// ==========================================
// 1. DATA MATRIX & COUNTRY MAP (Updated from Excel)
// ==========================================
const RATES = {
  "0.5": [421712, 431731, 456323, 503686, 574000, 601144, 688583, 906796],
  "1": [501864, 510062, 519170, 596590, 728000, 780390, 831583, 1114015],
  "1.5": [550697, 546093, 579000, 670000, 856000, 855000, 964000, 1274000],
  "2": [552688, 630000, 664000, 766000, 984000, 982000, 1103000, 1448000],
  "2.5": [790693, 927122, 942060, 1120314, 1112000, 1312511, 1426036, 1923953],
  "3": [792685, 930110, 945047, 1123470, 1221000, 1351539, 1429023, 1926941],
  "3.5": [794676, 933097, 948035, 1159282, 1331000, 1336303, 1467000, 1929928],
  "4": [829863, 975088, 990648, 1179625, 1440000, 1445000, 1584000, 2054000],
  "4.5": [831937, 1002000, 1064000, 1216000, 1550000, 1558000, 1700000, 2202000],
  "5": [857841, 1074000, 1143000, 1304000, 1659000, 1670000, 1817000, 2423613],
  "5.5": [1079485, 1379163, 1424180, 1544540, 1778000, 1886111, 1999637, 2703692],
  "6": [1136248, 1465105, 1516261, 1642132, 1897000, 2002624, 2116149, 2861034],
  "6.5": [1193010, 1509686, 1565453, 1739724, 2017000, 2119137, 2232662, 3018376],
  "7": [1249773, 1593336, 1655078, 1837316, 2136000, 2235650, 2349175, 3175718],
  "7.5": [1272600, 1676986, 1744703, 1934907, 2255000, 2352162, 2465687, 333060],
  "8": [1327888, 1714905, 1786683, 1979707, 2374000, 2468675, 2582200, 3490402],
  "8.5": [1383176, 1796383, 1873980, 2074764, 2493000, 2585188, 2698713, 3647744],
  "9": [1438464, 1877860, 1961278, 2169821, 2612000, 2701700, 2815226, 3805086],
  "9.5": [1493753, 1959338, 2048575, 2264878, 2732000, 2818213, 2931738, 3962427],
  "10": [1549041, 2040815, 2135872, 2359935, 2851000, 2934726, 3048251, 4119769],
  "10.5": [1556150, 2062891, 2159323, 2385277, 2974000, 3011091, 3164764, 4221816],
  "11": [1602475, 2136633, 2236847, 2470364, 3098000, 3126070, 3281276, 4378071],
  "11.5": [1648800, 2210376, 2314371, 2555451, 3221000, 3241050, 3397789, 4534325],
  "12": [1695125, 2284118, 2391895, 2640538, 3345000, 3356029, 3514302, 4690579],
  "12.5": [1741451, 2357860, 2469419, 2725626, 3468000, 3471009, 3630814, 4846834],
  "13": [1787776, 2431602, 2546943, 2810713, 3592000, 3585989, 3747327, 5003088],
  "13.5": [1834101, 2505344, 2624466, 2895800, 3716000, 3700968, 3863840, 5159343],
  "14": [1880426, 2579087, 2701990, 2980887, 3839000, 3815948, 3980352, 5315597],
  "14.5": [1926752, 2652829, 2779514, 3065974, 3963000, 3930928, 4096865, 5471851],
  "15": [1973077, 2726571, 2857038, 3151061, 4086000, 4045907, 4213378, 5628106],
  "15.5": [2019402, 2800313, 2934562, 3236149, 4210000, 4106849, 4329891, 5709238],
  "16": [2065727, 2874055, 3012086, 3321236, 4333000, 4220336, 4446403, 5863464],
  "16.5": [2112053, 2947798, 3089610, 3406323, 4457000, 4333822, 4562916, 6017689],
  "17": [2158378, 3021540, 3167133, 3491410, 4580000, 4447309, 4679429, 6171914],
  "17.5": [2204703, 3095282, 3244657, 3576497, 4704000, 4560795, 4795941, 6326139],
  "18": [2251028, 3169024, 3322181, 3661584, 4827000, 4674281, 4912454, 6480364],
  "18.5": [2297354, 3242767, 3399705, 3746672, 4951000, 4787768, 5028967, 6634589],
  "19": [2343679, 3316509, 3477229, 3831759, 5075000, 4901254, 5145479, 6788814],
  "19.5": [2390004, 3390251, 3554753, 3916846, 5198000, 5014741, 5261992, 6943039],
  "20": [2436329, 3463993, 3632277, 4001933, 5322000, 5128227, 5378505, 7097264],
  "20.5": [2388182, 3397376, 3563146, 3927476, 5436000, 5122000, 5474105, 7098377],
  "21": [2429169, 3457491, 3626904, 3999431, 5549000, 5210000, 5569705, 7229912],
  "21.5": [2470157, 3517605, 3690662, 4071386, 5663000, 5297000, 5665305, 7360072],
  "22": [2511144, 3577719, 3754419, 4143342, 5777000, 5385000, 5760905, 7491634],
  "22.5": [2552131, 3637834, 3818177, 4215297, 5891000, 5472000, 5856506, 7621817],
  "23": [2593118, 3697948, 3881935, 4287252, 6005000, 5560000, 5952106, 7753403],
  "23.5": [2634105, 3758063, 3945693, 4359207, 6119000, 5648000, 6047706, 7885002],
  "24": [2675092, 3818177, 4009450, 4431162, 6233000, 5735000, 6143306, 8015215],
  "24.5": [2716079, 3878292, 4073208, 4503117, 6347000, 5823000, 6238906, 8146836],
  "25": [2757066, 3938406, 4136966, 4575072, 6461000, 5910000, 6334506, 8277067],
  "25.5": [2990931, 4180381, 4380728, 4822937, 6673905, 6229948, 6504952, 8640548],
  "26": [3031541, 4239944, 4443901, 4894232, 6795681, 6314328, 6598168, 8767119],
  "26.5": [3072152, 4299506, 4507074, 4965527, 6917458, 6398709, 6691385, 8893690],
  "27": [3112763, 4359069, 4570246, 5036821, 7039234, 6483090, 6784602, 9020260],
  "27.5": [3153374, 4418632, 4633419, 5108116, 7161011, 6567470, 6877819, 9146831],
  "28": [3193985, 4478195, 4696592, 5179411, 7282787, 6651851, 6971035, 9273402],
  "28.5": [3234596, 4537758, 4759764, 5250706, 7404564, 6736231, 7064252, 9399973],
  "29": [3275207, 4597320, 4822937, 5322001, 7526341, 6820612, 7157469, 9526544],
  "29.5": [3315818, 4656883, 4886109, 5393295, 7648117, 6904993, 7250686, 9653115],
  "30": [3356429, 4716446, 4949282, 5464950, 7769894, 6989373, 7343902, 9779686]
};

const COUNTRIES = {
  "Afghanistan (AF)": 8, "Albania (AL)": 8, "Algeria (DZ)": 8, "Amenia (AM)": 8, "American Samoa (AS)": 8,
  "Andorra (AD)": 7, "Angola (AO)": 8, "Anguilla (AI)": 8, "Antigua (AG)": 8, "Argentina (AR)": 8,
  "Aruba (AW)": 8, "Australia (AU)": 4, "Austria (AT)": 7, "Azerbaijan (AZ)": 8, "Bahamas (BS)": 8,
  "Bahrain (BH)": 6, "Bangladesh (BD)": 6, "Barbados (BB)": 8, "Belarus (BY)": 8, "Belgium (BE)": 7,
  "Belize (BZ)": 8, "Benin (BJ)": 8, "Bermuda (BM)": 8, "Bhutan (BT)": 8, "Bolivia (BO)": 8,
  "Bonaire (XB)": 8, "Bosnia & Harzegovina (BA)": 8, "Botswana (BW)": 8, "Brazil (BR)": 8, "Brunei (BN)": 2,
  "Bulgaria (BG)": 7, "Burkina Faso (BF)": 8, "Burundi (BI)": 8, "Cambodia (KH)": 2, "Cameroon (CM)": 8,
  "Canada (CA)": 5, "Canary Islands,The (IC)": 8, "Cape Verde (CV)": 8, "Cayman Islands (KY)": 8, "Central African Rep (CF)": 8,
  "Chad (TD)": 8, "Chile (CL)": 8, "China (CN)*1": 3, "China (CN)*2": 4, "Cole D Invoice (CI)": 8,
  "Colombia (CO)": 8, "Comoros (KM)": 8, "Congo (CG)": 8, "Congo,DPR (CD)": 8, "Cook Islands (CK)": 8,
  "Costa Rica (CR)": 8, "Croatia (HR)": 7, "Cuba (CU)": 8, "Curacao (XC)": 8, "Cyprus (CY)": 7,
  "Czech Rep, The (GZ)": 7, "Denmark (DK)": 7, "Djibouti (DJ)": 8, "Dominica (DM)": 8, "Dominican Rep. (DO)": 8,
  "Ecuador (EC)": 8, "Egypt (EG)": 8, "El Salvador (SV)": 8, "Eritrea (ER)": 8, "Estonia (EE)": 7,
  "Eswatini (SZ)": 8, "Ethiopia (ET)": 8, "Faikland Islands (FK)": 8, "Faroe Islands (FO)": 8, "Fiji (FI)": 8,
  "Finland (FI)": 7, "France (FR)": 7, "French Guyana (GF)": 8, "Gabon (GA)": 8, "Gambia (GM)": 8,
  "Georgia (GE)": 8, "Germany (DE)": 7, "Ghana (GH)": 8, "Gibraltar (GI)": 8, "Greece (GR)": 7,
  "Greenland (GL)": 8, "Grenada (GD)": 8, "Guadelope (GP)": 8, "Guam (GU)": 8, "Guatemala (GU)": 8,
  "Guernsey (GG)": 8, "Guinea - Bissau (GW)": 8, "Guinea - Equatorial (GQ)": 8, "Guinea - Rep. (GN)": 8, "Guyana (British) (GY)": 8,
  "Haiti (HT)": 8, "Honduras (HN)": 8, "Hong Kong SAR China (HK)": 2, "Hungary (HU)": 7, "Iceland (IS)": 8,
  "India (in)": 6, "Iran (IR)": 8, "Iraq (IQ)": 8, "Ireland, Rep. Of (IE)": 7, "Israel (IL)": 8,
  "Italy (IT)": 7, "Jamaica (JM)": 8, "Japan (JP)": 3, "Jersey (JE)": 8, "Jordan (JO)": 6,
  "Kazakkhstan (KZ)": 8, "Kenya (KE)": 8, "Kiribati (KI)": 8, "Korea, D.P.R Of (KP)": 8, "Korea, Rep, Of (KR)": 4,
  "Kosovo (KV)": 8, "Kuwait (KW)": 6, "Kyrgyzstan (KG)": 8, "Lalvia (LV)": 7, "Laos (LA)": 2,
  "Lebanon (LB)": 8, "Lesotho (LS)": 8, "Liberia (LR)": 8, "Libya (LY)": 8, "Liechtenstein (LI)": 7,
  "Lithuania (LT)": 7, "Luxembourg (LU)": 7, "Macau SAR China (MO)": 2, "Madagascar (MG)": 8, "Malawi (MW)": 8,
  "Malaysia (MY)": 2, "Maldives (MV)": 6, "Mali (ML)": 8, "Malta (MT)": 7, "Marshall Islands (MH)": 8,
  "Martinique (MQ)": 8, "Mauritania (MR)": 8, "Mauritius (MU)": 8, "Mayotte (YT)": 8, "Mexico (MX)": 5,
  "Micronesia (FM)": 8, "Moldova, Rep, Of (MD)": 8, "Monaco (MC)": 7, "Mongalia (MN)": 8, "Monstserrat (MS)": 8,
  "Montenegro, Rep Of (ME)": 8, "Morocco (MA)": 8, "Mozambique (MZ)": 8, "Myanmar (MM)": 2, "Namibia (NA)": 8,
  "Nauru, Rep, Of (NR)": 8, "Nepal (NP)": 6, "Netherlands, The (NL)": 7, "Nevis (XN)": 8, "New Caledonia (NC)": 8,
  "New Zealand (NZ)": 4, "Nicaragua (NI)": 8, "Niger (NE)": 8, "Nigeria (NG)": 8, "Niue (NU)": 8,
  "North Macedonia (MK)": 8, "Northern Mariana Islands (MP)": 8, "Norway (NO)": 7, "Oman (OM)": 6, "Pakistan (PK)": 6,
  "Palau (PW)": 8, "Panama (PA)": 8, "Papua New Guinea (PG)": 4, "Paraguay (PY)": 8, "Peru (PE)": 8,
  "Philippines, The (PH)": 2, "Poland (PL)": 7, "Portugal (PT)": 7, "Puerto Rico (PR)": 8, "Qatar (QA)": 6,
  "Reunion, Island Of (RE)": 8, "Romania (RO)": 7, "Russian Federation (RU)": 8, "Rwanda (RW)": 8, "Saint Helena (SH)": 8,
  "Samoa (WS)": 8, "San Marino  (SM)": 7, "Sao Tome and Principe  (ST)": 8, "Saudi Arabia (SA)": 6, "Senegal (SN)": 8,
  "Serbia, Rep. Of (RS)": 8, "Seychelles (SC)": 8, "Sierra Leone (SL)": 8, "Singapore (SG)": 1, "Slovakia (SK)": 7,
  "Slovenia (SI)": 7, "Solomon Islands (SB)": 8, "Somalia (SO)": 8, "Somaliland, Rep Of (XS)": 8, "South Africa (ZA)": 7,
  "South Saudan (SS)": 8, "Spain (ES)": 7, "Sri Lanka (LK)": 6, "St. Barthelemy (XY)": 8, "St. Eustatius (XE)": 8,
  "St. Kitts (KN)": 8, "St. Lucia (LC)": 8, "St. Maarten (XM)": 8, "St. Vincent (VC)": 8, "Sudan (SD)": 8,
  "Suriname (SR)": 8, "Sweden (SE)": 7, "Switzerland (CH)": 7, "Syria (SY)": 8, "Tahiti (PF)": 8, "Taiwan (TW)": 4,
  "Tajikistan (TJ)": 8, "Tanzania (TZ)": 8, "Thailand (TH)": 2, "Timor-Leste (TL)": 1, "Togo (TG)": 8, "Tonga (TO)": 8,
  "Trinidad And Tabago (TT)": 8, "Tunisia (TN)": 8, "Turkey (TR)": 7, "Turkmenistan (TM)": 8, "Turks & Caicos (TC)": 8,
  "Tuvalu (TV)": 8, "USA (US)": 5, "Uganda (UG)": 8, "Ukraine (UA)": 8, "United Arab Emitayrs (AE)": 6,
  "United Kingdom (GB)": 7, "Uruguay (UY)": 8, "Uzbekistan (UZ)": 8, "Vanuatu (VU)": 8, "Vatica City (VA)": 7,
  "Venezuela (VE)": 8, "Vietnam (VN)": 2, "Virgin Island - US (VI)": 8, "Virgin Islands-British (VG)": 8, "Yemen, Rep. Of (YE)": 8,
  "Zambia (ZM)": 8, "Zimbabwe  (ZW)": 8
};

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
  
  const zoneNum = COUNTRIES[toDestination.value] || 8;
  document.getElementById('zoneClassificationVal').textContent = `Zone ${zoneNum}`;
}

// ==========================================
// 4. RATE ENGINE CALCULATIONS
// ==========================================
function calculateEstimate() {
  const targetCountry = toDestination.value;
  if (!targetCountry) return;

  const zoneIdx = COUNTRIES[targetCountry] - 1;
  const inputWeight = parseFloat(document.getElementById('weight').value) || 0.5;
  const lengthCm = parseFloat(document.getElementById('length').value) || 0;
  const widthCm = parseFloat(document.getElementById('width').value) || 0;
  const heightCm = parseFloat(document.getElementById('height').value) || 0;

  let volumetricWeight = 0;
  if (lengthCm > 0 && widthCm > 0 && heightCm > 0) {
    volumetricWeight = (lengthCm * widthCm * heightCm) / 5000;
  }

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

  if (document.getElementById('service').value === 'express') {
    freight *= 1.40;
  }

  if (remoteAreaCheck.checked) {
    surcharge += 450000;
  }

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

populateCountries();
updateRouteChip();
