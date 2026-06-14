/***************
 * CNV Automation Simple Import
 *
 * Flow:
 * Campaign mẫu ID + Bearer + Sheet
 * → duplicate campaign
 * → đọc nodeMap campaign mới
 * → update CONDITION / SEND_REWARD / SEND_ZNS
 ***************/

const CNV_BASE_URL = 'https://hub.cnvcdp.com/api';

const DATA_SHEET_NAME = 'Logic CNV-Tờ rơi';

const START_ROW = 2;

// A: Tên phần thưởng
// B: rewardId
// C: Hạn sử dụng
// D: handung
const COL_REWARD_NAME = 1;
const COL_REWARD_ID = 2;
const COL_END_TIME = 3;
const COL_HANDUNG = 4;

function onOpen(e) {
  menuCNV();
}

function menuCNV() {
  SpreadsheetApp.getUi()
    .createMenu('CNV Automation')
    .addItem('Tạo automation từ sheet', 'showCnvDialog')
    .addToUi();
}

function showCnvDialog() {
  const html = HtmlService
    .createHtmlOutputFromFile('Dialog')
    .setWidth(520)
    .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, 'Tạo automation CNV');
}

/**
 * Main.
 */
function runCnvSimpleImport(form) {
  const templateCampaignId = cleanText_(form.templateCampaignId);
  const bearerToken = normalizeBearerToken_(form.bearerToken);

  if (!templateCampaignId) {
    throw new Error('Chưa nhập Campaign mẫu ID.');
  }

  if (!bearerToken) {
    throw new Error('Chưa nhập Bearer token.');
  }

  const rows = readDataRows_();

  if (!rows.length) {
    throw new Error('Sheet không có dữ liệu hợp lệ.');
  }

  const auth = {
    bearerToken
  };

  /**
   * 1. Duplicate campaign mẫu.
   */
  const duplicateResponse = apiRequest_(
    'POST',
    `/campaign/${templateCampaignId}/duplicate`,
    auth,
    {}
  );

  const newCampaignId = extractCampaignId_(duplicateResponse);

  if (!newCampaignId) {
    throw new Error(
      'Đã duplicate nhưng không tìm thấy campaign ID mới trong response.'
    );
  }

  /**
   * 2. Đợi CNV tạo xong campaign copy.
   */
  Utilities.sleep(1200);

  /**
   * 3. Lấy nodeMap campaign mới.
   */
  const nodeMapResponse = apiRequest_(
    'GET',
    `/campaign/${newCampaignId}/node-map`,
    auth
  );

  const nodeMap = extractNodeMap_(nodeMapResponse);

  const branches = extractBranches_(nodeMap);
  const triggerNode = getTriggerNode_(nodeMap);
  const fallbackItemId = Number(triggerNode.id);

  if (!branches.length) {
    throw new Error(
      'Không tìm thấy flow dạng CONDITION → SEND_REWARD → SEND_ZNS trong campaign mới.'
    );
  }

  if (rows.length > branches.length) {
    throw new Error(
      'Sheet có ' + rows.length + ' dòng nhưng campaign mẫu chỉ có ' +
      branches.length + ' nhánh.\n\n' +
      'Hãy thêm nhánh trong campaign mẫu hoặc giảm số dòng sheet.'
    );
  }

  /**
   * 4. Update từng nhánh.
   */
  const results = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const branch = branches[i];

    const conditionResult = updateConditionNode_(
      branch.condition.id,
      row.rewardName,
      auth,
      fallbackItemId
    );

    const itemId = conditionResult.itemId;

    updateRewardNode_(
      branch.reward.id,
      row.rewardId,
      row.endTime,
      itemId,
      auth
    );

    updateZnsNode_(
      branch.zns.id,
      row.handung,
      itemId,
      auth
    );

    results.push({
      rowNumber: row.rowNumber,
      conditionNodeId: branch.condition.id,
      rewardNodeId: branch.reward.id,
      znsNodeId: branch.zns.id
    });
  }

  /**
   * 5. Nếu campaign mẫu có nhiều nhánh hơn số dòng sheet,
   * tự vô hiệu hoá các nhánh dư bằng condition không bao giờ match.
   */
  const disabledResults = [];

  for (let i = rows.length; i < branches.length; i++) {
    const branch = branches[i];

    updateConditionNode_(
      branch.condition.id,
      '__DISABLED__' + newCampaignId + '_' + (i + 1),
      auth,
      fallbackItemId
    );

    disabledResults.push({
      branchIndex: i + 1,
      conditionNodeId: branch.condition.id
    });
  }

  return {
    ok: true,
    templateCampaignId,
    newCampaignId,
    totalRows: rows.length,
    totalBranches: branches.length,
    disabledBranches: disabledResults.length,
    results,
    disabledResults
  };
}

/**
 * Read Sheet rows.
 */
function readDataRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME) || ss.getActiveSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) return [];

  const values = sheet
    .getRange(
      START_ROW,
      1,
      lastRow - START_ROW + 1,
      4
    )
    .getValues();

  return values
    .map((r, index) => {
      return {
        rowNumber: START_ROW + index,
        rewardName: cleanText_(r[COL_REWARD_NAME - 1]),
        rewardId: cleanRewardId_(r[COL_REWARD_ID - 1]),
        endTime: r[COL_END_TIME - 1],
        handung: cleanText_(r[COL_HANDUNG - 1])
      };
    })
    .filter(row => {
      return row.rewardName || row.rewardId || row.handung;
    })
    .map(row => {
      if (!row.rewardName) {
        throw new Error('Dòng ' + row.rowNumber + ' thiếu Tên phần thưởng ở cột A.');
      }

      if (!row.rewardId) {
        throw new Error(
          'Dòng ' + row.rowNumber +
          ' thiếu rewardId ở cột B. Cột B phải là số, ví dụ 21074.'
        );
      }

      if (!row.endTime) {
        throw new Error('Dòng ' + row.rowNumber + ' thiếu hạn sử dụng ở cột C.');
      }

      if (!row.handung) {
        throw new Error('Dòng ' + row.rowNumber + ' thiếu handung ở cột D.');
      }

      return row;
    });
}

/**
 * Update CONDITION.
 *
 * API:
 * PUT /api/condition/{nodeId}
 */
function updateConditionNode_(conditionNodeId, rewardName, auth, fallbackItemId) {
  let currentConfig = null;

  try {
    const current = apiRequest_(
      'GET',
      `/condition/${conditionNodeId}`,
      auth
    );

    currentConfig = extractConfig_(current);
  } catch (err) {
    currentConfig = null;
  }

  if (!currentConfig) {
    currentConfig = {
      stepFilters: [
        {
          conditionType: 'BASIC',
          advancedCondition: null,
          itemId: null,
          orFilters: [
            [
              {
                field: 'rewardName',
                operator: 'EQ',
                value: '',
                useFormula: null
              }
            ]
          ]
        }
      ],
      conditionType: 'BASIC',
      advancedCondition: null
    };
  }

  currentConfig.conditionType = currentConfig.conditionType || 'BASIC';
  currentConfig.advancedCondition = currentConfig.advancedCondition || null;

  if (!currentConfig.stepFilters || !currentConfig.stepFilters.length) {
    currentConfig.stepFilters = [
      {
        conditionType: 'BASIC',
        advancedCondition: null,
        itemId: null,
        orFilters: [
          [
            {
              field: 'rewardName',
              operator: 'EQ',
              value: '',
              useFormula: null
            }
          ]
        ]
      }
    ];
  }

  const stepFilter = currentConfig.stepFilters[0];

  if (!stepFilter.itemId && fallbackItemId) {
    stepFilter.itemId = Number(fallbackItemId);
  }

  stepFilter.conditionType = stepFilter.conditionType || 'BASIC';
  stepFilter.advancedCondition = stepFilter.advancedCondition || null;

  if (!stepFilter.orFilters || !stepFilter.orFilters.length) {
    stepFilter.orFilters = [[]];
  }

  if (!stepFilter.orFilters[0] || !stepFilter.orFilters[0].length) {
    stepFilter.orFilters[0] = [
      {
        field: 'rewardName',
        operator: 'EQ',
        value: '',
        useFormula: null
      }
    ];
  }

  const filter = stepFilter.orFilters[0][0];

  filter.field = 'rewardName';
  filter.operator = 'EQ';
  filter.value = rewardName;
  filter.useFormula = null;

  if (!stepFilter.itemId) {
    throw new Error(
      'Không tìm thấy itemId trong CONDITION node ' +
      conditionNodeId +
      '. Hãy mở lại node condition mẫu và lưu lại một lần.'
    );
  }

  apiRequest_(
    'PUT',
    `/condition/${conditionNodeId}`,
    auth,
    {
      config: currentConfig
    }
  );

  return {
    itemId: stepFilter.itemId
  };
}

/**
 * Update SEND_REWARD.
 *
 * API:
 * PUT /api/send-reward-action/{nodeId}
 */
function updateRewardNode_(rewardNodeId, rewardId, endTimeValue, itemId, auth) {
  let currentConfig = null;

  try {
    const current = apiRequest_(
      'GET',
      `/send-reward-action/${rewardNodeId}`,
      auth
    );

    currentConfig = extractConfig_(current);
  } catch (err) {
    currentConfig = null;
  }

  if (!currentConfig) {
    currentConfig = {};
  }

  currentConfig.rewardId = Number(rewardId);
  currentConfig.slot = currentConfig.slot || 1;

  currentConfig.customerConfig = itemId;
  currentConfig.phoneNumberConfig = 'customerPhone';
  currentConfig.fullItemPhoneNumberConfig = `{${itemId}|customerPhone}`;

  /**
   * Theo payload bạn gửi:
   * isFollowExpiredAt vẫn là true nhưng endTime có giá trị.
   */
  currentConfig.isFollowExpiredAt = true;
  currentConfig.isIssueCodeNow = false;
  currentConfig.endTime = convertSheetDateToCnvIso_(endTimeValue);

  apiRequest_(
    'PUT',
    `/send-reward-action/${rewardNodeId}`,
    auth,
    {
      config: currentConfig
    }
  );
}

/**
 * Update SEND_ZNS_TEMPLATE.
 *
 * API:
 * PUT /api/send-zns-action/{nodeId}
 */
function updateZnsNode_(znsNodeId, handung, itemId, auth) {
  let currentConfig = null;

  try {
    const current = apiRequest_(
      'GET',
      `/send-zns-action/${znsNodeId}`,
      auth
    );

    currentConfig = extractConfig_(current);
  } catch (err) {
    currentConfig = null;
  }

  if (!currentConfig) {
    currentConfig = {
      oaId: 'c19bfd7f-e2ca-4af0-ae79-96800843d179',
      templateId: '08deb72e-5f9f-4b2a-89d9-c625136336a6',
      sendType: 'AUTO',
      supportSendByUid: true,
      acceptedRegulationAt: new Date().toISOString(),
      acceptedRegulationBy: null,
      templateFields: {}
    };
  }

  currentConfig.customerConfig = String(itemId);
  currentConfig.customerIdConfig = 'customerId';
  currentConfig.fullCustomerIdConfig = `{${itemId}|customerId}`;

  currentConfig.templateFields = currentConfig.templateFields || {};

  currentConfig.templateFields.customer_name = `{${itemId}|customerName}`;
  currentConfig.templateFields.phone = `{${itemId}|customerPhone}`;
  currentConfig.templateFields.phan_thuong = `{${itemId}|rewardId}`;
  currentConfig.templateFields.han_dung = handung;

  apiRequest_(
    'PUT',
    `/send-zns-action/${znsNodeId}`,
    auth,
    {
      config: currentConfig
    }
  );
}

/**
 * API helper.
 */
function apiRequest_(method, path, auth, body) {
  const url = CNV_BASE_URL + path;

  const options = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Accept: 'application/json, text/plain, */*',
      Authorization: auth.bearerToken
    }
  };

  if (body !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const text = response.getContentText();

  if (status >= 200 && status < 300) {
    if (!text) return null;

    try {
      return JSON.parse(text);
    } catch (err) {
      return text;
    }
  }

  throw new Error(
    method + ' ' + path + ' lỗi HTTP ' + status + '\n\n' +
    text.slice(0, 1000)
  );
}

/**
 * Extract helpers.
 */
function extractCampaignId_(response) {
  if (!response) return null;

  if (response.data && response.data.id) return response.data.id;
  if (response.data && response.data.campaignId) return response.data.campaignId;

  if (response.id) return response.id;
  if (response.campaignId) return response.campaignId;

  return null;
}

function extractNodeMap_(response) {
  if (!response) {
    throw new Error('Response nodeMap rỗng.');
  }

  if (response.data && response.data.nodeMap) return response.data.nodeMap;
  if (response.nodeMap) return response.nodeMap;
  if (response.nodes && response.connections) return response;

  throw new Error('Không tìm thấy nodeMap trong response.');
}

function extractConfig_(response) {
  if (!response) return null;

  const data = response.data || response;

  if (data.configData) return data.configData;
  if (data.config) return data.config;

  if (data.actionData && data.actionData.config) {
    return data.actionData.config;
  }

  if (data.actionData && typeof data.actionData === 'object') {
    return data.actionData;
  }

  return null;
}

/**
 * Extract flow branches:
 * TRIGGER
 * → CONDITION
 * → SEND_REWARD
 * → SEND_ZNS_TEMPLATE
 * otherwise → next CONDITION
 */
function extractBranches_(nodeMap) {
  const nodes = nodeMap.nodes || [];
  const connections = nodeMap.connections || [];

  const nodeById = {};

  nodes.forEach(node => {
    nodeById[String(node.id)] = node;
  });

  const trigger = nodes.find(node => {
    return node.itemType === 'TRIGGER';
  });

  if (!trigger) {
    throw new Error('Không tìm thấy TRIGGER node.');
  }

  const branches = [];

  let currentCondition = findNextNodeBySocket_(
    trigger.id,
    'then',
    connections,
    nodeById
  );

  const visited = {};

  while (currentCondition && !visited[String(currentCondition.id)]) {
    visited[String(currentCondition.id)] = true;

    if (currentCondition.itemType !== 'CONDITION') {
      break;
    }

    const reward = findNextNodeBySocket_(
      currentCondition.id,
      'then',
      connections,
      nodeById
    );

    if (!reward || reward.subType !== 'SEND_REWARD') {
      break;
    }

    const zns = findNextNodeBySocket_(
      reward.id,
      'then',
      connections,
      nodeById
    );

    if (!zns || zns.subType !== 'SEND_ZNS_TEMPLATE') {
      break;
    }

    branches.push({
      condition: currentCondition,
      reward: reward,
      zns: zns
    });

    currentCondition = findNextNodeBySocket_(
      currentCondition.id,
      'otherwise',
      connections,
      nodeById
    );
  }

  return branches;
}


function getTriggerNode_(nodeMap) {
  const nodes = nodeMap.nodes || [];

  const trigger = nodes.find(node => {
    return node.itemType === 'TRIGGER';
  });

  if (!trigger) {
    throw new Error('Không tìm thấy TRIGGER node trong nodeMap.');
  }

  return trigger;
}

function findNextNodeBySocket_(sourceNodeId, socketName, connections, nodeById) {
  const sourceNode = nodeById[String(sourceNodeId)];

  if (!sourceNode) return null;

  const outputId = getOutputIdBySocket_(sourceNode, socketName);

  const conn = connections.find(c => {
    return (
      String(c.source) === String(sourceNodeId) &&
      String(c.sourceOutput) === String(outputId)
    );
  });

  if (!conn) return null;

  return nodeById[String(conn.target)] || null;
}

function getOutputIdBySocket_(node, socketName) {
  if (!node || !node.outputs) {
    throw new Error('Node ' + node.id + ' không có outputs.');
  }

  const outputKeys = Object.keys(node.outputs);

  for (let i = 0; i < outputKeys.length; i++) {
    const output = node.outputs[outputKeys[i]];

    if (
      output &&
      output.socket &&
      String(output.socket.name) === String(socketName)
    ) {
      return output.id;
    }
  }

  throw new Error(
    'Không tìm thấy output socket "' +
    socketName +
    '" trong node ' +
    node.id
  );
}

/**
 * Date converter.
 *
 * Sheet:
 * 26/06/2026
 *
 * CNV:
 * 2026-06-25T17:00:00.000Z
 */
function convertSheetDateToCnvIso_(value) {
  let year;
  let month;
  let day;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    year = Number(Utilities.formatDate(value, 'Asia/Ho_Chi_Minh', 'yyyy'));
    month = Number(Utilities.formatDate(value, 'Asia/Ho_Chi_Minh', 'MM'));
    day = Number(Utilities.formatDate(value, 'Asia/Ho_Chi_Minh', 'dd'));
  } else {
    const text = cleanText_(value);

    const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

    if (!match) {
      throw new Error(
        'Ngày không đúng định dạng dd/MM/yyyy: ' + text
      );
    }

    day = Number(match[1]);
    month = Number(match[2]);
    year = Number(match[3]);
  }

  /**
   * 00:00 ngày dd/MM/yyyy theo giờ VN = 17:00 ngày trước đó UTC.
   */
  const date = new Date(Date.UTC(year, month - 1, day, -7, 0, 0, 0));

  return date.toISOString();
}

/**
 * Utils.
 */
function cleanText_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function cleanRewardId_(value) {
  const text = cleanText_(value);

  if (!text) return '';

  const match = text.match(/\d+/);

  if (!match) return '';

  return Number(match[0]);
}

function normalizeBearerToken_(token) {
  const clean = cleanText_(token);

  if (!clean) return '';

  if (clean.toLowerCase().startsWith('bearer ')) {
    return clean;
  }

  return 'Bearer ' + clean;
}