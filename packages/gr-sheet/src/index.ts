import type {
  RegisterItem,
  RegisterConfiguration,
  ItemClassConfiguration,
  InternalItemReference,
} from '@riboseinc/paneron-registry-kit/types';
import type { Predicate } from '@riboseinc/paneron-registry-kit/proposals/objectChangeset.js';
import type { CommonGRItemData } from '@riboseinc/paneron-extension-geodetic-registry/classes/common.js';
import type { Extent } from '@riboseinc/paneron-extension-geodetic-registry/classes/extent.js';
import type { EllipsoidData } from '@riboseinc/paneron-extension-geodetic-registry/classes/ellipsoid.js';
import type { DatumData, GeodeticDatumData } from '@riboseinc/paneron-extension-geodetic-registry/classes/datum.js';
import type { TransformationParameter, TransformationData } from '@riboseinc/paneron-extension-geodetic-registry/classes/transformation.js';
import type { ConversionParameter, ConversionData } from '@riboseinc/paneron-extension-geodetic-registry/classes/conversion.js';
import type { CoordinateSystemData } from '@riboseinc/paneron-extension-geodetic-registry/classes/coordinate-systems.js';
import type { CoordinateSystemAxisData } from '@riboseinc/paneron-extension-geodetic-registry/classes/coordinate-sys-axis.js';
import type { CoordinateOpMethod } from '@riboseinc/paneron-extension-geodetic-registry/classes/coordinate-op-method.js';
import type { UoMData } from '@riboseinc/paneron-extension-geodetic-registry/classes/unit-of-measurement.js';
import type { NonCompoundCRSData } from '@riboseinc/paneron-extension-geodetic-registry/classes/crs.js';

import xlsx, { readSheetNames, type Row } from 'read-excel-file';

import type { FileConvertor } from '../../common/src/convertors/index.js';
import { teeAsync } from '../../common/src/util.js';


// Duplicating from GR extension b/c we cannot import it due to bad packaging
export const ParameterType = {
  FILE: 'parameter file name',
  MEASURE: 'measure (w/ UoM)',
  INTEGER_VALUE: 'integer value',
} as const;


export interface GRSheetConvertor
extends FileConvertor<
  ParsedSheetItem,
  GRItem<CommonGRItemData>,
  GRConfig> {}


// TODO: This should be possible to obtain using `typeof itemClassConfiguration`
// with GR extension’s itemClassConfiguration, but TS somehow loses
// type information about register item payloads.
export interface GRConfig extends RegisterConfiguration<{
  "coordinate-ops--conversion": ItemClassConfiguration<ConversionData>,
  "coordinate-ops--transformation": ItemClassConfiguration<TransformationData>,
  "coordinate-sys-axis": ItemClassConfiguration<CoordinateSystemAxisData>,
  "coordinate-system": ItemClassConfiguration<CoordinateSystemData>,
  "datums--engineering": ItemClassConfiguration<DatumData>,
}> {
  subregisters: undefined,
}

// /**
//  * Maps
//  * sheet ID aliases (initial part of e.g. CA#, CS#)
//  * to
//  * item class IDs
//  *
//  */
// const ItemClassSheetIDPrefixes:
// Record<string, (parsedRow: Record<string, string>) => keyof GRConfig["itemClassConfiguration"]> = {
//   CA: () => 'coordinate-sys-axis',
//   CS: () => 'coordinate-system',
//   CR: (row) => `crs--${row.type!.split(' ')[0]!.toLowerCase()}`,
// } as const;

export const Sheets = {
  EXTENTS: `Geo_Extent(GE#)`,
  CITATIONS: 'Source_Citation(CI#)',
  OPERATION_PARAM_VALUES: 'ParamVal(PV#)',
  COORDINATE_OP_PARAMS: 'OpParam(OP#)',
  COORDINATE_OP_METHODS: 'OpMethod(OM#)',

  TRANSFORMATIONS: 'Coord_Trans(CT#)',
  CONVERSIONS: 'Coord_Conv(CC#)',
  COMPOUND_CRS: 'CompCRS(CM#)',
  NON_COMPOUND_CRS: 'CRS(CR#)',

  COORDINATE_SYSTEMS: 'CoordSys(CS#)',
  COORDINATE_SYSTEM_AXES: 'CSAxis(CA#)',
  UOM: 'UoM(UM#)',

  DATUMS: 'Datum(CD#)',

  ELLIPSOIDS: 'Ellips(EL#)',
} as const;
type SheetName = typeof Sheets[keyof typeof Sheets];
function isSheetName(val: string): val is SheetName {
  return Object.values(Sheets).indexOf(val as typeof Sheets[keyof typeof Sheets]) >= 0;
}


/** Extracts the “CS” part from a sheet name like “CoordSys(CS#)”. */
function getSheetIDAlias(sheetName: string): string {
  return sheetName.split('(')[1]!.slice(0, 2);
}
/** Converts “CS” to full sheet name like “CoordSys(CS#)”. */
function getSheetName(alias: string): SheetName {
  const sheetName = Object.values(Sheets).find(sheetName => getSheetIDAlias(sheetName) === alias);
  if (sheetName && isSheetName(sheetName)) {
    return sheetName;
  } else {
    console.warn("Possible aliases", Object.values(Sheets).map(sheetName => getSheetIDAlias(sheetName)));
    throw new Error(`Unable to get sheet name from ${alias} (got ${sheetName})`);
  }
}


/** Intermediate item obtained after deserializing a sheet into JS. */
interface ParsedSheetItem {
  /** Sheet name */
  sheet: SupportedSheetName;

  /** Raw row data as from parse-excel-file */
  rowRaw: Row;

  /** Row data somewhat parsed */
  rowParsed: Record<string, string>;
}


/** An output item representing some GR item (no register data). */
interface GRItem<T extends CommonGRItemData> {
  /** GR item class ID (e.g., transformation) */
  itemRef: InternalItemReference;
  itemData: T;
}


export default function getConvertor(): GRSheetConvertor {
  return {
    label: "ISO GR Sheet Convertor (v8)",
    inputDescription: "One or more TC 211 GR v8 spreadsheet files, in XLSX format, containing proposed additions",
    parseInput,
    generateItems,
    generateRegisterItems,
  };
}

const parseInput: GRSheetConvertor["parseInput"] =
async function * parseSpreadsheetFiles(fileGenerator, opts) {
  // We can assume there to be multiple files.
  for await (const file of fileGenerator()) {
    function fileProgress(msg?: string) {
      const prefix = `Processing file ${file.name}`;
      opts?.onProgress?.(msg ? `${prefix}: ${msg}` : prefix);
    }

    try {
      const sheetNames = await readSheetNames(file.blob);
      for (const sheet of sheetNames/*.filter(sheet => SKIP_SHEETS.indexOf(sheet) < 0)*/) {
        const rows: Row[] = await xlsx(file.blob, { sheet });
        for (const [idx, row] of rows.entries()) {
          if (idx < 3) {
            // Skip header rows
            continue;
          } else if (row[0] === null) {
            // Skip empty rows
            continue;
          } else if (!isSheetName(sheet)) {
            // Skip unrecognized sheets
            fileProgress(`Skipping unrecognized sheet ${sheet}`);
            continue;
          } else if (!isSupportedSheetName(sheet)) {
            // Skip unsupported sheets
            fileProgress(`WARNING: Skipping sheet ${sheet}: not yet supported`);
            continue;
          }
          const processor = SupportedSheets[sheet];
          yield {
            sheet,
            rowRaw: row,
            rowParsed:
              processor.fields.
                map((fname, idx) => ({
                  [fname as string]:
                    // Avoid nulls
                    row[idx] !== null
                      ? `${row[idx]}`
                      : ''
                })).
                reduce((prev, curr) => ({ ...prev, ...curr })),
          };
        }
      }
    } catch (e) {
      fileProgress(`Error: ${(e as any).toString?.() ?? "No error information available"}`);
    }
  }
}


const generateItems: GRSheetConvertor["generateItems"] =
async function * generateGRItems(parsedSheetItems, opts) {
  const [stream1, stream2] = teeAsync(parsedSheetItems());

  opts?.onProgress?.("Caching items");
  const cache = await cacheItems(stream1);

  const idMap: TemporaryIDMap = {};
  let availableID = -1;

  const getOrCreateIdentifiers = function (rowParsed: Record<string, string>): { ref: InternalItemReference, identifier: number } {
    console.debug("Generating item IDs", rowParsed.sheetID);

    if (!rowParsed.sheetID) {
      throw new Error("No sheetID in parsed row, cannot get or create identifiers");
    }
    if (!idMap[rowParsed.sheetID]) {
      const sheetName = getSheetName(rowParsed.sheetID.slice(0, 2));
      const processor = SupportedSheets[sheetName];

      if (isRegisterItemProcessor(processor)) {
        const classID = processor.getClassID(
          rowParsed as Record<Exclude<(typeof processor)["fields"][number], null>, string>
        );
        const itemID = crypto.randomUUID();
        const itemRef = { classID, itemID };
        const identifier = availableID;
        availableID = availableID - 1;
        idMap[rowParsed.sheetID] = { ref: itemRef, identifier };

      } else {
        throw new Error(`Unable to create a reference for a non-register item procesor (${sheetName})`);
      }
    }
    return idMap[rowParsed.sheetID]!;
  }

  const resolveReference = function (cellContents: string, mode: Predicate["mode"]): Predicate | InternalItemReference | string {
    console.debug("Resolving ref", cellContents);

    const itemID = extractItemID(cellContents);
    try {
      const item = resolveRelated(itemID);
      if (item) {
        const ref = getOrCreateIdentifiers(item).ref;
        return mode === 'generic' ? ref : ref.itemID;
      } else {
        console.warn(`Referenced item ‘${itemID}’ cannot be found in this proposal, got:`, item);
      }
    } catch (e) {
      console.warn(`Referenced item ‘${itemID}’ cannot be found in this proposal`, cellContents, e);
      return predicate(
        makePredicateQuery(itemID),
        mode,
      );
    }
    opts?.onProgress?.(`Referenced item ${itemID} cannot be found in this proposal`);
    throw new Error(`Unable to resolve reference, ${itemID}`);
  }

  const isCellNull = function (cellContents: string): boolean {
    return cellContents.trim().length === 0 || cellContents.trim().toLowerCase() === 'none';
  }

  const resolveNullableReference = function (cellContents: string, mode: Predicate["mode"]): ReturnType<typeof resolveReference> | null {
    if (isCellNull(cellContents)) {
      return null;
    }
    return resolveReference(cellContents, mode);
  }

  const resolveRelated = function resolveRelated(cellContents: string) {

    //const itemID = sheetItemID.split(' ')[0];
    const sheetItemID = extractItemID(cellContents);
    console.debug("Resolving item", sheetItemID, cellContents);
    const sheetName = getSheetName(sheetItemID.slice(0, 2));
    if (cache[sheetName]) {
      const parsedRow = cache[sheetName]![sheetItemID];
      if (parsedRow) {
        return parsedRow;
      } else {
        console.warn("Cache for sheet", sheetName, cache[sheetName]);
        throw new Error(`Cannot resolve related item ${sheetItemID}`);
      }
    } else {
      console.warn("ALL CACHE", cache);
      throw new Error(`Cannot resolve item from sheet ${sheetName} (no data for that sheet)`);
    }
  }

  const constructItem = function constructItem(parsedRow: Record<string, string>): unknown {
    console.debug("Constructing item", parsedRow.sheetID);

    const sheetName = getSheetName(parsedRow.sheetID!.slice(0, 2));
    const processor = SupportedSheets[sheetName];
    const pr = parsedRow as Record<Exclude<(typeof processor)["fields"][number], null>, string>
    if (isRegisterItemProcessor(processor)) {
      return processor.toRegisterItem(pr, resolveAndConstruct, resolveReference, opts);
    } else if (isBasicSheetItemProcessor(processor)) {
      return processor.toItem(pr, resolveAndConstruct, resolveNullableReference, opts);
    } else {
      throw new Error("Unknown processor");
    }
  }

  const resolveAndConstruct = function resolveAndConstruct(sheetItemID: string): unknown {
    console.debug("Resolving and constructing item", sheetItemID);

    return constructItem(resolveRelated(sheetItemID));
  }

  for await (const sheetItem of stream2) {
    // Process actual items
    const processor = SupportedSheets[sheetItem.sheet];
    if (isRegisterItemProcessor(processor)) {
      let parsedItem: Omit<CommonGRItemData, 'identifier'>;

      if (!sheetItem.rowParsed.sheetID) {
        throw new Error(`Sheet ID column is missing in parsed row data, ${sheetItem.rowParsed}`);
      }

      const { ref: itemRef, identifier } = getOrCreateIdentifiers(sheetItem.rowParsed);

      const pr = sheetItem.rowParsed as Record<Exclude<(typeof processor)["fields"][number], null>, string>

      try {
        opts?.onProgress?.(`Handling GR item ${sheetItem.rowParsed.sheetID}`);
        parsedItem = processor.toRegisterItem(
          pr,
          resolveAndConstruct,
          resolveReference,
        );

      } catch (e) {
        console.warn("Unable to transform sheet row to item", sheetItem.sheet, sheetItem.rowParsed, e);
        opts?.onProgress?.(`${sheetItem.sheet}/${sheetItem.rowRaw[0]}: error processing register item: ${String(e)}`);
        continue;
        //throw e;
      }

      console.debug("Processed", sheetItem, "into", parsedItem);
      opts?.onProgress?.(`Creating GR item ${itemRef.classID}`);

      yield {
        itemData: {
          identifier,
          ...parsedItem,
        },
        itemRef,
      };

    } else {
      console.debug("Skipping entity", sheetItem);
      opts?.onProgress?.(`${sheetItem.sheet}/${sheetItem.rowRaw[0]}: not outputting a register item`);
    }
  }
}


const generateRegisterItems: GRSheetConvertor["generateRegisterItems"] =
async function * generateRegisteredGRItems(grItems, opts) {
  // Current timestamp
  const dateAccepted = new Date();
  let idx = 0;
  for await (const { itemRef, itemData } of grItems) {
    const item: RegisterItem<any> = {
      id: itemRef.itemID,
      data: itemData,
      dateAccepted,
      status: 'valid',
    };
    yield {
      [itemRef.classID]: item,
    }
    opts?.onProgress?.(`Outputting as register items: #${idx + 1} (UUID “${itemRef.itemID}”)`);
  }
}


/** Spec for a particular sheet to be processed. */
interface BaseSheetItemProcessor<F extends string, I> {
  /**
   * Which column maps to which field in the object produced.
   * `null` means column data is ignored.
   */
  fields: (F | null)[];
  toItem: (
    /** Row parsed into fields based on `fields` spec given. */
    item: Record<F, string>,
    /** Retrieve entire item data. For non-register items from other sheets, e.g. extents. */
    getSheetItem: (sheetItemID: string) => unknown | undefined,
    /**
     * Resolve link, either to predicate to resolve a preexisting register item at import time
     * or as `InternalItemReference` referencing item being added in the same proposal.
     */
    resolveReference:
      (rawCellContents: string, mode: Predicate["mode"]) =>
        Predicate | InternalItemReference | string | null,
    opts?: { onProgress?: (msg: string) => void },
  ) => I;
}
function isBasicSheetItemProcessor(val: unknown): val is BaseSheetItemProcessor<any, any> {
  return val && val.hasOwnProperty('toItem') ? true : false;
}
/** Fields common for all sheets. */
type CommonColumns = 'sheetID'

function makeProcessor<F extends string, I>
(p: BaseSheetItemProcessor<Exclude<F, CommonColumns>, I>):
BaseSheetItemProcessor<F, I> {
  (p as BaseSheetItemProcessor<F | CommonColumns, I>).fields = [
    'sheetID',
    ...p.fields,
  ];
  return p;
}


/** Spec for a sheet to be processed into individual GR items. */
interface RegisteredItemProcessor<F extends string, RI>
extends Omit<BaseSheetItemProcessor<F, RI>, 'toItem'> {
  getClassID: (item: Record<F, string>) => string;
  toRegisterItem: (
    /** Row parsed into fields based on `fields` spec given. */
    item: Record<F, string>,
    /** Retrieve entire item data. For non-register items from other sheets, e.g. extents. */
    getSheetItem: (sheetItemID: string) => unknown | undefined,
    /**
     * Resolve link, either to predicate to resolve a preexisting register item at import time
     * or as `InternalItemReference` referencing item being added in the same proposal.
     */
    resolveReference:
      (rawCellContents: string, mode: Predicate["mode"]) =>
        Predicate | InternalItemReference | string | null,
    opts?: { onProgress?: (msg: string) => void },
  ) => RI;
}
function isRegisterItemProcessor(val: unknown): val is RegisteredItemProcessor<any, any> {
  return val && val.hasOwnProperty('getClassID') ? true : false;
}

type RequiredRegisterItemFields = keyof CommonGRItemData;

type ProposalRelatedColumns =
  'justification'
| 'registerManagerNotes'
| 'controlBodyNotes'
| 'check';

type GuaranteedRegisterItemSheetColumns =
  CommonColumns
| RequiredRegisterItemFields
| ProposalRelatedColumns;

/** Fields that change position from sheet to sheet. */
type VariableRegisterItemFields =
  'informationSources'
| 'remarks';

function makeItemProcessor<F extends string, RI>
(p: RegisteredItemProcessor<
  Exclude<F, 'citation' | 'citations' | 'informationSource' | Exclude<GuaranteedRegisterItemSheetColumns, VariableRegisterItemFields>>,
  Omit<RI, keyof CommonGRItemData>
>): RegisteredItemProcessor<
  F | GuaranteedRegisterItemSheetColumns,
  RI & Omit<CommonGRItemData, 'identifier'>
> {
  const pp = p as unknown as RegisteredItemProcessor<
    F | GuaranteedRegisterItemSheetColumns,
    RI & Omit<CommonGRItemData, 'identifier'>
  >;

  pp.fields = [
    'sheetID',
    'name',
    'aliases',
    ...p.fields,
    'justification',
    'registerManagerNotes',
    'controlBodyNotes',
    null,
    'check',
  ];

  const originalParser = p.toRegisterItem;

  const parseRegisterItemWrapped: (typeof pp)["toRegisterItem"] =
  function parseRegisterItemWrapped (parsedRow, resolveRelated, resolveReference) {
    const commonData: Omit<CommonGRItemData, 'identifier'> = {
      name: parsedRow.name.trim() || '',
      aliases: (parsedRow.aliases ?? '').trim() !== ''
        ? parsedRow.aliases.split(';').map((a: string) => a.trim())
        : [],
      informationSources: (parsedRow.informationSources ?? '').trim() !== ''
        ? parsedRow.informationSources.split(';').
            map((cid: string) => resolveRelated(extractItemID(cid)))
        : [],
      remarks: (parsedRow.remarks ?? '').trim() || '',
    }
    const data = originalParser(parsedRow, resolveRelated, resolveReference);
    const result = {
      ...commonData,
      ...data,
    } as RI & Omit<CommonGRItemData, 'identifier'>;
    return result;
  }

  // XXX don’t cast
  pp.toRegisterItem = parseRegisterItemWrapped;

  return pp;
}

/** Helper for typing return values of toRegisterItem */
type Item<Payload, PredicateFields extends keyof Payload> = Omit<UsePredicates<Payload, PredicateFields>, keyof CommonGRItemData>

const SupportedSheets = {
  [Sheets.TRANSFORMATIONS]: makeItemProcessor({
    fields: [
      'sourceCRS', 'targetCRS',
      null,  // <- Operation type (always “transformation”)
      'scope', 'remarks', 'coordinateOperationMethod', 'extent', 'params', 'operationVersion', 'accuracy', 'informationSources'],
    getClassID: function () {
      return 'coordinate-ops--transformation';
    },
    toRegisterItem: function toTransformation(item, resolveRelated, resolveReference, opts) {
      const extentRef = resolveReference(item.extent, 'id') as string;
      //const itemData: Omit<ReplaceKeys<
      //  UsePredicates<TransformationData, 'sourceCRS' | 'targetCRS'>,
      //  'accuracy',
      //  UsePredicates<TransformationData["accuracy"], 'unitOfMeasurement'>
      //>, 'identifier'> =
      const parameters = item.params.split(';').
        map(p => p.trim()).
        filter(p => p !== '').
        map(paramSheetID => resolveRelated(extractItemID(paramSheetID)) as TransformationParameter).
        map(({ type, value, unitOfMeasurement, parameter, fileCitation }) => {
          const param: TransformationParameter = {
            type,
            fileCitation,
            value,
            parameter,
            unitOfMeasurement,
          };
          return param;
        });
      const data: Item<TransformationData, 'sourceCRS' | 'targetCRS' | 'coordOperationMethod'> = {
        operationVersion: item.operationVersion,
        coordOperationMethod: resolveReference(item.coordinateOperationMethod, 'id'),
        scope: item.scope,
        // TODO: Not required, UoM is always metre.
        accuracy: parseValueWithUoM(item.accuracy) as unknown as TransformationData["accuracy"],
        sourceCRS: resolveReference(item.sourceCRS, 'generic'),
        targetCRS: resolveReference(item.targetCRS, 'generic'),
        extentRef,
        parameters,
      };
      return data;
    },
  }),
  [Sheets.CONVERSIONS]: makeItemProcessor({
    fields: [null, 'scope', 'remarks', 'coordinateOperationMethod', 'extent', 'parameters', 'informationSources'],
    getClassID: () => 'coordinate-ops--conversion',
    toRegisterItem: function toConversion(item, resolveRelated, resolveReference, opts) {
      const extentRef = resolveReference(item.extent, 'id') as string;
      if (!extentRef) {
        throw new Error("No extent!");
      }
      const parameters = item.parameters.split(';').
        map(p => p.trim()).
        filter(p => p !== '').
        map(paramSheetID => resolveRelated(extractItemID(paramSheetID)) as TransformationParameter).
        map(({ type, value, unitOfMeasurement, parameter }) => {
          if (type === ParameterType.FILE) {
            opts?.onProgress?.("ERROR: “Reference File” parameters are not supported on Conversions");
            //throw new Error("“Reference File” parameters are not supported on Conversions");
          }
          const param: ConversionParameter = {
            //name,
            value,
            parameter,
            unitOfMeasurement,
          };
          return param;
        });
      const result: Item<ConversionData, 'coordinateOperationMethod'> = {
        coordinateOperationMethod: resolveReference(item.coordinateOperationMethod, 'id'),
        scope: item.scope,
        parameters,
        extentRef,
      };
      return result;
    },
  }),
  [Sheets.COMPOUND_CRS]: makeItemProcessor({
    fields: ['scope', 'remarks', 'horizontalCRS', 'verticalCRS', 'extent', 'informationSources'],
    getClassID: () => 'crs--compound',
    toRegisterItem: function toCompoundCRS(item, resolveRelated, resolveReference) {
      const extentRef = resolveReference(item.extent, 'id');
      if (!extentRef) {
        throw new Error("No extent!");
      }
      return {
        remarks: item.remarks,
        horizontalCRS: resolveReference(item.horizontalCRS, 'generic'),
        verticalCRS: resolveReference(item.verticalCRS, 'generic'),
        extentRef,
        informationSources: [],
        scope: item.scope,
      };
    },
  }),
  [Sheets.NON_COMPOUND_CRS]: makeItemProcessor({
    fields: ['scope', 'remarks', 'type', 'datum', 'coordinateSystem', 'baseCRS', 'operation', 'extent', 'informationSources'],
    getClassID: (row) => `crs--${row.type.split(' ')[0]!.toLowerCase()}`,
    toRegisterItem: function toNonCompoundCRS(item, resolveRelated, resolveReference) {
      const extentRef = resolveReference(item.extent, 'id') as string;

      const baseCRS = item.baseCRS.trim() !== ''
        ? resolveReference(item.baseCRS, 'generic')
        : null;

      const operation = item.operation.trim() !== ''
        ? resolveReference(item.operation, 'generic')
        : null;

      type NonCompoundCRSPredicateFieldNames = 'coordinateSystem' | 'baseCRS' | 'operation';
      type SharedData = Item<Omit<NonCompoundCRSData, 'datum'>, NonCompoundCRSPredicateFieldNames>;
      const shared: SharedData = {
        scope: item.scope,
        coordinateSystem: resolveReference(item.coordinateSystem, 'generic'),
        baseCRS,
        operation,
        extentRef,
      };

      switch (item.type) {
        case 'Vertical CRS':
          return {
            ...shared,
            datum: resolveReference(item.datum, 'id'),
          };
        case 'Geodetic CRS':
          return {
            ...shared,
            datum: resolveReference(item.datum, 'id'),
          };
        case 'Projected CRS':
          return shared;
        // Doesn’t seem allowed by the spreadsheet
        // case 'Engineering CRS':
        //   itemType = 'crs--engineering';
        default:
          throw new Error(`Unknown CRS type: ${item.type}`);
      }
    },
  }),
  [Sheets.COORDINATE_SYSTEMS]: makeItemProcessor({
    fields: ['type', 'remarks', 'coordinateSystemAxes', 'informationSources'],
    getClassID: (row) => `coordinate-sys--${row.type.replace(" Coordinate System", '').trim().toLowerCase()}`,
    toRegisterItem: function toCoordinateSystem(item, _resolveRelated, resolveReference) {
      const axes = item.coordinateSystemAxes.split(';').
        map(a => a.trim()).
        map(axisID => resolveReference(axisID, 'id'));

      return {
        coordinateSystemAxes: axes,
      };
    },
  }),
  [Sheets.COORDINATE_SYSTEM_AXES]: makeItemProcessor({
    fields: ['remarks', 'abbreviation', 'orientation', 'unitOfMeasurement', 'minimumValue', 'maximumValue', 'rangeMeaning', 'informationSources'],
    getClassID: () => 'coordinate-sys-axis',
    toRegisterItem: function toCoordinateSystemAxis(item, _resolveRelated, resolveReference) {
      return {
        remarks: item.remarks,
        orientation: item.orientation,
        abbreviation: item.abbreviation,
        unitOfMeasurement: resolveReference(item.unitOfMeasurement, 'id'),
      };
    },
  }),
  [Sheets.UOM]: makeItemProcessor({
    fields: ['remarks', 'baseUnit', 'numerator', 'denominator', 'measureType', 'symbol', 'informationSources'],
    getClassID: () => 'unit-of-measurement',
    toRegisterItem: function toUoM(item, _resolveRelated, resolveReference) {
      const c: Omit<UoMData, 'baseUnit' | 'identifier' | keyof CommonGRItemData> & { baseUnit?: Predicate | string } = {
        symbol: item.symbol,
        numerator: item.numerator.trim() !== '' ? parseInt(item.numerator, 10) : null,
        denominator: item.denominator.trim() !== '' ? parseInt(item.denominator, 10) : null,
        // XXX
        measureType: item.measureType as any,
      };
      if (item.baseUnit?.trim?.() != '') {
        c.baseUnit = resolveReference(item.baseUnit, 'id') as string | Predicate;
      }
      return c;
    },
  }),
  [Sheets.COORDINATE_OP_PARAMS]: makeItemProcessor({
    fields: ['remarks', 'minimumOccurs', 'informationSources'],
    getClassID: () => 'coordinate-op-parameter',
    toRegisterItem: function parseCoordinateOpParam({ minimumOccurs }) {
      return {
        minimumOccurs: parseInt(minimumOccurs.trim(), 10),
      };
    },
  }),
  [Sheets.COORDINATE_OP_METHODS]: makeItemProcessor({
    fields: ['remarks', 'parameters', 'formula', 'formulaCitation'],
    getClassID: () => 'coordinate-op-method',
    toRegisterItem: function parseCoordinateOpMethod({ parameters, formula, formulaCitation }, resolveRelated, resolveReference) {
      const item: Omit<UsePredicateLists<CoordinateOpMethod, 'parameters'>, 'identifier' | keyof CommonGRItemData> = {
        parameters: parameters.trim() !== ''
          ? parameters.split(';').map(paramUUID => resolveReference(paramUUID, 'id'))
          : [],
      };
      if (formulaCitation.trim()) {
        item.formulaCitation = resolveRelated(formulaCitation);
      } else if (formula.trim()) {
        item.formula = formula;
      }
      return item;
    },
  }),
  [Sheets.ELLIPSOIDS]: makeItemProcessor({
    fields: ['remarks', 'semiMajorAxis', 'globalUoM', 'isSphere', 'inverseFlattening', 'semiMinorAxis', 'informationSources'],
    getClassID: () => 'ellipsoid',
    toRegisterItem: function parseEllipsoid(item, resolveRelated, resolveReference) {
      const uom = resolveReference(item.globalUoM, 'id') as string | Predicate | null;
      if (uom === null) {
        throw new Error("Unit of measure cannot be null");
      }
      console.info({ uom });
      const isSphereString = item.isSphere.toLowerCase();
      const isSphere = isSphereString === 'false'
        ? false
        : isSphereString === 'TRUE'
          ? true
          : undefined;
      if (isSphere === undefined) {
        throw new Error(`isSphere is not correctly defined: ${item.isSphere}`);
      }
      const d: Omit<UsePredicates<EllipsoidData, 'semiMajorAxisUoM' | 'semiMinorAxisUoM' | 'inverseFlatteningUoM'>, keyof CommonGRItemData> = {
        semiMajorAxisUoM: uom,
        semiMinorAxisUoM: uom,
        inverseFlatteningUoM: uom,

        semiMajorAxis: parseFloat(item.semiMajorAxis),
        semiMinorAxis: parseFloat(item.semiMinorAxis),
        inverseFlattening: parseFloat(item.inverseFlattening),
        isSphere,
      };
      return d;
    },
  }),
  [Sheets.DATUMS]: makeItemProcessor({
    fields: ['type', 'scope', 'remarks', 'originDescription', 'ellipsoid', 'primeMeridian', 'releaseDate', 'coordinateReferenceEpoch', 'extent', 'informationSources'],
    getClassID: ({ type }) => (type === 'Vertical Datum' ? 'datums--vertical' : 'datums--geodetic'),
    toRegisterItem: function parseDatum({ scope, originDescription, releaseDate, ...item }, resolveRelated, resolveReference) {
      const extentRef = resolveReference(item.extent, 'id') as Extent | undefined;
      if (!extentRef) {
        throw new Error("No extent!");
      }
      const sharedData: Item<DatumData, never> = {
        scope,
        originDescription,
        releaseDate,
        extentRef,
        coordinateReferenceEpoch: item.coordinateReferenceEpoch.trim() || null,
      } as const;
      if (/geodetic(?: *datum *)?$/i.test(item.type)) {
        const d: Omit<UsePredicates<GeodeticDatumData, 'ellipsoid' | 'primeMeridian'>, keyof CommonGRItemData> = {
          ...sharedData,
          ellipsoid: resolveReference(item.ellipsoid, 'id'),
          primeMeridian: resolveReference(item.primeMeridian, 'id'),
        };
        return d;
      } else {
        if (item.ellipsoid || item.primeMeridian) {
          throw new Error(`Ellipsoid and prime meridian are not recognized as properties of a ${item.type}`);
        }
        return sharedData;
      }
    },
  }),
  [Sheets.OPERATION_PARAM_VALUES]: makeProcessor({
    fields: [
      'parameter',
      null, // <- Link to transformation -- useless? We link from transformation to here instead
      'type', 'value', 'unitOfMeasurement',
      null, // <- UoM name -- redundant?
      'fileRef', 'citation',
    ],
    toItem: function parseTransformationParam({ parameter, type, value, unitOfMeasurement, fileRef, citation }, resolveRelated, resolveReference, opts) {
      const c: ReplaceKeys<UsePredicates<TransformationParameter, 'parameter'>, 'unitOfMeasurement', string | Predicate | null> = {
        parameter: resolveReference(parameter, 'id') as string | Predicate,
        type: type === "Reference File"
          ? ParameterType.FILE
          : type === "Integer"
          ? ParameterType.INTEGER_VALUE
          : ParameterType.MEASURE,
        unitOfMeasurement: ["Reference File", 'Integer'].includes(type)
          ? null
          : resolveReference(unitOfMeasurement, 'id') as string | Predicate | null,
        value: type === "Reference File"
          ? fileRef
          : value,
        //name: '', // XXX: name seems unused
        fileCitation: null,
      };
      if (citation.trim() !== '') {
        try {
          c.fileCitation = resolveRelated(citation);
        } catch (e) {
          opts?.onProgress?.(`ERROR: Failed to resolve op. parameter file reference citation based on “${citation}”`);
          console.error("Failed to resolve related", e);
          c.fileCitation = citation;
        }
      }
      return c;
    },
  }),
  [Sheets.EXTENTS]: makeItemProcessor({
    fields: ['description', 's', 'w', 'n', 'e', 'polygon', 'startDate', 'finishDate'],
    getClassID: () => 'extent',
    toRegisterItem: function parseExtent ({ description, s, w, n, e }) {
      return {
        extent: {
          name: description,
          s,
          w,
          n,
          e,
        },
      };
    },
  }),
  [Sheets.CITATIONS]: makeProcessor({
    fields: ['title', 'alternateTitles', 'author', 'publisher', 'publicationDate', 'revisionDate', 'edition', 'editionDate', 'seriesName', 'seriesIssueID', 'seriesPage', 'doi', 'otherDetails', 'uri'],
    toItem: function parseCitation ({ title, publicationDate, doi, revisionDate, edition, author, publisher, otherDetails, seriesName, seriesIssueID, seriesPage, uri }) {
      return {
        // XXX
        title,
        publicationDate,
        revisionDate,
        edition,
        author,
        seriesName,
        seriesIssueID,
        seriesPage,
        otherDetails,
        uri,
        doi,
        //alternateTitles: item.alternateTitles.split(';').map(t => t.trim()),
        publisher,
      };
    },
  }),
} as const;
type SupportedSheetName = keyof typeof SupportedSheets;
function isSupportedSheetName(val: string): val is SupportedSheetName {
  return SupportedSheets[val as SupportedSheetName] !== undefined;
}


type ReplaceKeys<T, Keys extends keyof T, WithType> = Omit<T, Keys> & { [key in Keys]: WithType };
type UsePredicates<T, Keys extends keyof T> = ReplaceKeys<T, Keys, Predicate | InternalItemReference | string | null>;
type UsePredicateLists<T, Keys extends keyof T> = ReplaceKeys<T, Keys, (Predicate | InternalItemReference | string)[]>;


/**
 * Extracts referenced item ID from raw cell contents.
 */
const REFERENCE_SEPARATOR = ' - ';
function extractItemID(cellValue: string): string {
  let id: string;
  if (cellValue.indexOf(REFERENCE_SEPARATOR) > 1) {
    const parts: string[] = cellValue.split(REFERENCE_SEPARATOR);
    if (parts.length < 2) {
      throw new Error(`Unable to extract a reference from ${cellValue}`);
    }

    id = parts[0]!;
  } else {
    id = cellValue.trim();
  }

  // Validate?
  // try {
  //   getSheetName(id.slice(0, 2))
  // }

  return id;
}


/** Maps sheet IDs, like CM1, to item refs and temporary GR identifiers. */
type TemporaryIDMap = Record<string, { ref: InternalItemReference, identifier: number }>;

/**
 * Returns predicate if preexisting item’s numeric ID is found in raw cell data;
 * otherwise assumes it’s cross-referencing an item added in the same proposal,
 * so looks up or generates a new temporary ID and UUID for the item in question
 * and (depending on given predicate mode)
 * returns either InternalItemReference or item ID string.
 */
function makePredicateQuery(
  sheetItemID: string,
): string {
  // Preexisting items are referenced by numerical identifiers in the sheet.
  let idNum: number | undefined = undefined;
  try {
    idNum = parseInt(sheetItemID, 10);
  } catch (e) {
    idNum = undefined;
  }
  idNum = typeof idNum === 'number' && (idNum > 0 || idNum < 0)
    ? idNum
    : undefined;

  // idNum can be a NaN. XD
  if (idNum !== undefined) {
    return  `data.identifier === ${idNum}`;
  } else {
    throw new Error(`Identifier ‘${sheetItemID}’ is unparseable or invalid`);
  }
}

function predicate(query: string, mode: Predicate["mode"]): Predicate {
  return {
    __isPredicate: true,
    mode,
    predicate: query,
  };
}

/**
 * Separates value with UoM into a numerical value
 * and UoM pointer.
 */
function parseValueWithUoM(raw: string): { value: number, unitOfMeasurement: Predicate } {
  let uomAlias: any;
  let uomRaw = raw.slice(raw.length - 1);
  try {
    uomAlias = parseInt(uomRaw, 10);
    // If uomAlias parses as a number, it’s not real.
    uomAlias = 'm';
  } catch (e) {
    uomAlias = uomRaw;
  }
  const value = parseFloat(raw.endsWith(uomAlias)
    ? raw.substring(0, raw.length - 1)
    : raw);
  return {
    value,
    unitOfMeasurement: {
      __isPredicate: true,
      mode: 'id',
      predicate: `data.aliases?.indexOf("${uomAlias}") >= 0`,
    },
  };
}


/** Indexes row data by first column (ID), groups by sheet. */
type CachedItems = Record<SheetName, Record<string, Record<string, string>>>;

async function cacheItems(
  items: AsyncGenerator<ParsedSheetItem, void, undefined>,
) {
  const cache: Partial<CachedItems> = {};

  for await (const item of items) {
    cache[item.sheet] ??= {};
    cache[item.sheet]![item.rowRaw[0]] = item.rowParsed;
  }
  return cache;
}
