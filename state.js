// -----------------------------
// State
// -----------------------------
const DEFAULT_GROUP_DESCS = [
  'Bolt Components',
  'Bolts Finished Goods',
  'Buy - ins',
  'Civil Buy - ins',
  'Cable Bolt',
  'Cable Components',
  'Cable Raw Material',
  'Civil - Westconnex',
  'Drill Bits & Rods',
  'Friction-Lok',
  'J Bar Raw Material',
  'Low Profile Bolt',
  'Mesh',
  'Plastic Bolts',
  'Plastic Plate',
  'Plate Feed - Steel',
  'Plates - Steel',
  'Raw Material',
  'Rebar Raw Material',
  'Resin Injection',
  'Smooth Rod Raw Materials',
  'Steel Mesh',
  'Tensioners,Footpumps & Com',
  'Washers',
  'Work In Process',
];

const state = {
  audit: {
    useDefaultMatchRule: true,
    useDescTitleRule: false,
    useDescTitleTunableRule: false,
    tunableMinSharedTokens: 4,
    tunableMinJaccard: 0.50,

    stockMatchCol: 'Part Code',
    vaultMatchCol: 'Stock Number',

    stockMatchIdx: -1,
    stockGroupIdx: -1,
    stockDescIdx: -1,
    stockMoveDateIdx: -1,

    vaultMatchIdx: -1,
    vaultTypeIdx: -1,
    vaultStateIdx: -1,
    vaultTitleIdx: -1,

    vaultIndex: new Map(),
    titleTokenIndex: new Map(),
    titleToEntries: new Map(),

    counts: {
      released: 0,
      unreleased: 0,
      pdf: 0,
      modelled: 0,
      folder: 0,
      missing: 0,
      totalConsidered: 0,
    },

    review: {
      items: [],
      cursor: 0,
      decisions: new Map(), // key -> 'approved' | 'flagged'
    },
  },

  stock: {
    fileName: null,
    headers: [],
    rawRows: [],
    rows: [],
    groupCounts: new Map(),
    groupSelected: new Map(),
    dateFilter: {
      mode: 'none',
      start: null,
      end: null,
    },
  },

  vault: {
    fileName: null,
    headers: [],
    rows: [],
  },
};
