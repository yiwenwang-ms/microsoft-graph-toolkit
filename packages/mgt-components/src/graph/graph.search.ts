/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, CacheItem, CacheService, CacheStore } from '@microsoft/mgt-element';
import { schemas } from './cacheStores';

export interface SuggestionItem extends CacheItem {
  entity: string;
  referenceId: string;
}

/**
 * Object to be stored in cache representing individual people
 */
export interface SuggestionPeople extends SuggestionItem {
  /**
   * json representing a person stored as string
   */
  displayName?: string;
  personImage?: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  imAddress?: string;
}

/**
 * Object to be stored in cache representing individual people
 */
export interface SuggestionFile extends SuggestionItem {
  /**
   * json representing a person stored as string
   */
  name: string;
  addtionalInformation?: string;
  DateModified?: string;
  FileExtension?: string;
  FileType?: string;
  Query?: string;
  FileSize?: number;
  AccessUrl?: string;
  id?: string;
}

/**
 * Object to be stored in cache representing individual people
 */
export interface SuggestionQuery extends SuggestionItem {
  /**
   * json representing a person stored as string
   */
  query: string;
}

/**
 * Object to be stored in cache representing individual people
 */
export interface Suggestions extends CacheItem {
  /**
   * json representing a person stored as string
   */
  fileSuggestions: SuggestionFile[];
  querySuggestions: SuggestionQuery[];
  peopleSuggestions: SuggestionPeople[];
  otherSuggestion1?: any[];
  otherSuggestion2?: any[];
}

export interface SuggestionEntityConfig {
  maxCount: number;
}

export interface SuggestionConfig {
  configMap: Map<String, SuggestionEntityConfig>;
  queryString: string;
}

/**
 * Stores results of queries (multiple people returned)
 */
interface CachePeopleQuery extends CacheItem {
  /**
   * max number of results the query asks for
   */
  maxResults?: number;
  /**
   * list of people returned by query (might be less than max results!)
   */
  results?: string[];
}

const getIsPeopleCacheEnabled = (): boolean =>
  CacheService.config.suggestions.isEnabled && CacheService.config.isEnabled;

/**
 * Defines the expiration time
 */
const getSuggestionInvalidationTime = (): number =>
  CacheService.config.suggestions.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

const getMockData = (): Suggestions => {
  var mockFileSuggestions: SuggestionFile[] = [
    {
      entity: 'File',
      name: 'File A.docx',
      addtionalInformation: 'Append Description on Main Description Main Description Main Description Main Desc',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.1',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'docx',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    },
    {
      entity: 'File',
      name: 'File B.xlsx',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.2',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'xlsx',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    },
    {
      entity: 'File',
      name: 'File C.pptx',
      addtionalInformation: 'C Append Description on Main Description Main Description Main Description Main Desc',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.3',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'pptx',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    },
    {
      entity: 'File',
      name: 'File D.txt',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.4',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'txt',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    }
  ];

  var mockQuerySuggestions: SuggestionQuery[] = [
    {
      entity: 'Text',
      query: 'Query A',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.5'
    },
    {
      entity: 'Text',
      query:
        'any one of the vector terms added to form a vector sum or resultant / a coordinate of a vectoreither member of an ordered pair of numbers',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.6'
    },
    {
      entity: 'Text',
      query: 'Query C',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.7'
    },
    {
      entity: 'Text',
      query: 'Query D',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.8'
    },
    {
      entity: 'Text',
      query: 'Query E',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.9'
    }
  ];

  var mockPeopleSuggestions: SuggestionPeople[] = [
    {
      entity: 'People',
      displayName: 'Isaiah Langer',
      jobTitle: 'Web Marketing Manager',
      //personImage: 'http://image14.m1905.cn/uploadfile/2018/1107/20181107104420301408_watermark.jpg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.10',
      imAddress: 'sip:isaiahl@m365x214355.onmicrosoft.com'
    },
    {
      entity: 'People',
      displayName: 'Megan Bowen',
      jobTitle: 'Web Marketing Manager',
      //personImage: 'http://image14.m1905.cn/uploadfile/2018/1107/20181107104420301408_watermark.jpg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.11',
      imAddress: 'sip:meganb@m365x214355.onmicrosoft.com'
    },
    {
      entity: 'People',
      displayName: 'Alex Wilber',
      jobTitle: 'Web Marketing Manager',
      //personImage: 'http://5b0988e595225.cdn.sohucs.com/images/20171130/658b9f5ea4394831b90e3f65d3cd83b6.jpeg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.12'
    },
    {
      entity: 'People',
      displayName: 'bob@tailspin.com',
      //personImage: 'http://image14.m1905.cn/uploadfile/2018/1107/20181107104420301408_watermark.jpg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.13'
    },
    {
      entity: 'People',
      displayName: 'Lynne Robbins',
      personImage: '',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.14',
      imAddress: 'sip:lynner@m365x214355.onmicrosoft.com'
    }
  ];

  var otherSuggestion1 = [
    {
      entity: 'Sample1',
      referenceId: '1111111111111110'
    },
    {
      entity: 'Sample1',
      referenceId: '1111111111111111'
    },
    {
      entity: 'Sample1',
      referenceId: '1111111111111112'
    }
  ];

  var otherSuggestion2 = [
    {
      entity: 'Sample2',
      referenceId: '1111111111111120'
    },
    {
      entity: 'Sample2',
      referenceId: '1111111111111121'
    },
    {
      entity: 'Sample2',
      referenceId: '1111111111111122'
    }
  ];

  var mockSuggestions: Suggestions = {
    fileSuggestions: mockFileSuggestions,
    querySuggestions: mockQuerySuggestions,
    peopleSuggestions: mockPeopleSuggestions,
    otherSuggestion1: otherSuggestion1,
    otherSuggestion2: otherSuggestion2
  };
  return mockSuggestions;
};

/**
 * async promise to the Graph for Suggestions, by default, it will request the most frequent contacts for the signed in user.
 *
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getSuggestions(graph: IGraph, queryConfig: SuggestionConfig): Promise<Map<string, any[]>> {
  // mock action
  var mockSuggestions = getMockData();
  for (var key in mockSuggestions) {
    var suggestion = mockSuggestions[key];
    var entityType = suggestion[0].entity.toLowerCase();
    var config = queryConfig.configMap.get(entityType);
    if (config != null && config != undefined) {
      var maxCount = config.maxCount;
      mockSuggestions[key] = mockSuggestions[key]
        .filter(item => {
          if (item.entity == 'Text') {
            return item.query.toLowerCase().startsWith(queryConfig.queryString.toLowerCase());
          }

          if (item.entity == 'People') {
            return item.displayName.toLowerCase().startsWith(queryConfig.queryString.toLowerCase());
          }

          if (item.entity == 'File') {
            return item.name.toLowerCase().startsWith(queryConfig.queryString.toLowerCase());
          }

          return true;
        })
        .slice(0, maxCount);
    } else {
      mockSuggestions[key] = [];
    }
  }

  var resultMap = new Map();
  for (var key in mockSuggestions) {
    if (mockSuggestions[key].length > 0) {
      resultMap.set(mockSuggestions[key][0].entity.toLowerCase(), mockSuggestions[key]);
    }
  }
  return resultMap;
}
