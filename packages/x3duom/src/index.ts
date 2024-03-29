import type { LocalizedConceptData } from '@riboseinc/paneron-extension-glossarist/classes/localizedConcept/LocalizedConceptData.js';
import type { Designation, Expression } from '@riboseinc/paneron-extension-glossarist/models/concepts.js';
import type { FileConvertor } from '../../common/src/convertors/index.js';

import {
  type GlossaryRegisterConfig,
  type LinkParser,
  processLinks,
  asRegisterItemsWithInferredUniversal,
} from '../../common/src/glossary.js';


/** Item obtained from processing an X3D UOM XML file. */
interface IntermediateItem {
  id: string;

  /** <enumeration> */
  el: Element;

  /** Properties of a designation that cannot be read from XML element itself. */
  designationProperties: DesignationStub,
}


type DesignationStub =
    Pick<Designation, 'normative_status'>
  & Partial<Omit<Designation, 'designation'>>;


export interface X3DUOMConvertor
extends FileConvertor<
  IntermediateItem,
  LocalizedConceptData,
  GlossaryRegisterConfig> {}


export default function getConvertor(): X3DUOMConvertor {
  return {
    label: "X3D UOM XML",
    inputDescription: "An XML file, or a directory with XML files, containing terms in X3D UOM format",
    parseInput,
    generateItems,
    generateRegisterItems,
  };
}


const decoder = new TextDecoder('utf-8');
const parser = new DOMParser();


const parseInput: X3DUOMConvertor["parseInput"] =
async function * parseInput(fileGenerator, opts) {
  for await (const file of fileGenerator()) {

    function fileProgress(msg?: string) {
      const prefix = `Processing file ${file.name}`;
      opts?.onProgress?.(msg ? `${prefix}: ${msg}` : prefix);
    }

    const rawXML = decoder.decode(file.blob);
    fileProgress("Decoded file into string");

    try {
      const items = convertX3D(rawXML, fileProgress);
      for await (const item of items) {
        yield item;
      }
    } catch (e) {
      fileProgress(`Error: ${(e as any).toString?.() ?? "No error information available"}`);
    }
  }
}


const parseLinks: LinkParser =
function parseLinks(text, { forMatchingDesignation }) {
  const [newText, ] =  extractHashtags(text, function handleTag(_, parsedTag) {
    const _tag = parsedTag.toLowerCase();
    return forMatchingDesignation(t => t.designation.replace(' ', '').toLowerCase() === _tag);
  });
  return newText;
}


const generateItems: X3DUOMConvertor["generateItems"] =
async function * generateConcepts(itemGenerator, opts) {
  let idx = 1;

  for await (const item of itemGenerator()) {
    function itemProgress(msg?: string) {
      const prefix = `Parsing <${item.el.localName}> #${idx}: ${item.id}`;
      opts?.onProgress?.(msg ? `${prefix}: ${msg}` : prefix);
    }
    itemProgress(`Attributes ${JSON.stringify(item.designationProperties)}`);
    try {
      yield await parseLocalizedConcept(item, itemProgress);
    } catch (e) {
      itemProgress(`Error: ${(e as any)?.toString?.() ?? 'No error information available'}`);
    }
  }
}


const generateRegisterItems: X3DUOMConvertor["generateRegisterItems"] =
async function * generateGlossaryRegisterItems(itemGenerator, opts) {
  console.info("LINKS?", opts?.urnNamespace);
  const stream = opts?.urnNamespace
    ? processLinks(parseLinks, asRegisterItemsWithInferredUniversal(itemGenerator, opts), {
        linkURNPrefix: opts.urnNamespace,
        onProgress: opts.onProgress
          ? function (msg) { opts!.onProgress!(`Process links: ${msg}`); }
          : undefined,
      })
    : asRegisterItemsWithInferredUniversal(itemGenerator, opts);
  yield * stream;

  //if (opts?.urnNamespace) {
  //  return processLinks(parseLinks, asRegisterItems(itemGenerator), {
  //    linkURNPrefix: opts.urnNamespace,
  //  });
  //} else {
  //  return asRegisterItems(itemGenerator);
  //}
}


const convertX3D = async function * (
  xmlString: string,
  onProgress?: (msg: string) => void,
) {
  const doc = parser.parseFromString(xmlString, 'text/xml');
  const acronymContainers = doc.getElementsByName('acronymChoices');
  const glossaryContainers = doc.getElementsByName('glossaryChoices');
  if (acronymContainers.length < 1 && glossaryContainers.length < 1) {
    throw new Error("Specified file contained neither acronymChoices nor glossaryChoices");
  }
  yield * readSimpleType(
    acronymContainers,
    (msg) => onProgress?.(`Acronyms: ${msg}`),
    true,
  );
  yield * readSimpleType(
    glossaryContainers,
    (msg) => onProgress?.(`Glossary: ${msg}`),
  );
}


function * readSimpleType(
  simpleTypeEls: NodeListOf<HTMLElement>,
  onProgress?: (msg: string) => void,
  isAbbreviation?: boolean,
) {
  const containerEntries = [...simpleTypeEls.entries()];
  const total = containerEntries.
    map(([, el]) => el.children.length).
    reduce((acc, curr) => acc + curr, 0);
  for (const [containerIdx, container] of containerEntries) {
    const children = [...container.children].entries();
    for (const [enumIdx, maybeEnumEl] of children) {
      const decimalIdx = parseFloat(`${containerIdx + 1}.${enumIdx + 1}`)

      const val = maybeEnumEl.getAttribute('value');
      if (val) {
        onProgress?.(`<${maybeEnumEl.localName}> ${decimalIdx} or ${total}`);

        const expression: Pick<Expression, 'isAbbreviation'> = {}
        if (isAbbreviation) {
          expression.isAbbreviation = true;
        }

        const item: IntermediateItem = {
          id: val,
          el: maybeEnumEl,
          designationProperties: {
            normative_status: 'preferred',
            type: 'expression',
            ...expression,
          },
        };
        yield item;

      } else {
        onProgress?.(`Error: Empty value attribute`);
      }
    }
  }
}


/**
 * Matches hashtags and returns them as a list with leading hash character
 * and any trailing punctuation stripped.
 */
function extractHashtags(
  /** Text with possibly some hashtags. */
  text: string,

  /**
   * The function will be invoked with raw hashtag
   * (complete with hash character and any trailing punctuation)
   * and parsed hashtag (no hash or trailing punctuation).
   * If it returns a string, it will be used to replace hashtag occurrence.
   * Allows to optionally replace every hashtag with something else.
   */
    onHashtag?: (rawTag: string, parsedTag: string) => string | undefined,
): [text: string, hahstags: Set<string>] {
  const parts = text.split(' ');
  const tags = new Set<string>();
  for (const [idx, part] of parts.entries()) {
    if (part.startsWith('#')) {
      const tag = part;
      const parsed = tag.slice().replace(/[^\w\s\']|_$/g, "");
      const replacement = onHashtag?.(tag, parsed);
      if (replacement) {
        parts.splice(idx, 1, replacement);
      }
      tags.add(parsed);
    }
  }
  return [parts.join(' '), tags];
}


async function parseLocalizedConcept(
  item: IntermediateItem,

  onProgress: ((msg: string) => void) | undefined,
): Promise<LocalizedConceptData> {
  if (item.el.localName === 'enumeration') {
    const definition = item.el.getAttribute('appinfo');
    const designation = item.el.getAttribute('alias') ?? item.el.getAttribute('value');
    const link = item.el.getAttribute('documentation');
    if (definition?.trim() && designation?.trim()) {
      return {
        language_code: 'eng',
        terms: [{ ...item.designationProperties, designation } as Designation], // TODO: Avoid cast
        definition: [{ content: definition }],
        notes: [],
        examples: [],
        authoritativeSource: link ? [{ link }] : [],
      };
    } else {
      throw new Error("Element is missing appinfo and value, required to extract definition and designation.");
    }
  } else {
    throw new Error("Unexpected element type");
  }
}
