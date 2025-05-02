# !/usr/bin/env python

"""list import/expot filters from registry
"""

"""
  (C) Copyright 2020 Shojiro Fushimi, all rights reserved

  Redistribution and use in source and binary forms, with or
  without modification, are permitted provided that the following
  conditions are met:

   - Redistributions of source code must retain the above
    copyright notice, this list of conditions and the following
    disclaimer.

   - Redistributions in binary form must reproduce the above
    copyright notice, this list of conditions and the following
    disclaimer in the documentation and/or other materials
    provided with the distribution.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDER ``AS IS''
  AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
  FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT
  SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
  CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
  PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
  OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
  THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
  TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT
  OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY
  OF SUCH DAMAGE.
"""

import sys
from argparse import ArgumentParser
from re import compile as re_compile
from functools import reduce
import csv
import xml.etree.ElementTree

DEFAULT_FIELDS = ['DocumentService', 'UIName', 'Flags']
DEFAULT_FLAGS = ['IMPORT', 'EXPORT', 'DEFAULT', 'PREFERRED']
DEFAULT_KEY_FIELD = 'DocumentService'

DEFAULT_TYPE_FIELDS = ['Extensions', 'MediaType']

NS_RE = re_compile(r'\{(.*?)\}')

def component_data_dict(file_name: str, component_name: str, node_name: str):
    tree = xml.etree.ElementTree.parse(file_name)
    root = tree.getroot()
    namespace_mo = NS_RE.match(root.tag)
    if namespace_mo is None:
        return {}
    namespace = {'oor': namespace_mo.group(1)}
    data_dict = {}
    for component_data in root.findall(f'.//oor:component-data[@oor:name="{component_name}"]', namespace):
        component_node = component_data.find(f'node[@oor:name="{node_name}"]', namespace)
        if component_node is None:
            continue
        oor_name = f'{{{namespace["oor"]}}}name'
        for data_node in component_node.findall('node'):
            name = data_node.attrib.get(oor_name)
            if name is None:
                continue
            prop_dict = {}
            for prop in data_node.findall('prop'):
                prop_name = prop.attrib.get(oor_name)
                if prop_name is None:
                    continue
                value = prop.find('value')
                if value is not None:
                    value_text = value.text
                else:
                    value_text = None
                prop_dict[prop_name] = value_text
            data_dict[name] = prop_dict
    return data_dict

def main(argv):
    aparser = ArgumentParser()
    aparser.add_argument('files', nargs='*')
    aparser.add_argument('--out', '-o', dest='out', action='store', default=None)
    aparser.add_argument('--fields', dest='fields', action='store', default=','.join(DEFAULT_FIELDS))
    aparser.add_argument('--flags', dest='flags', action='store', default=','.join(DEFAULT_FLAGS))
    aparser.add_argument('--all-fields', dest='all_fields', action='store_true', default=False)
    aparser.add_argument('--all-flags', dest='all_flags', action='store_true', default=False)
    aparser.add_argument('--key-field', dest='key_field', action='store', default=DEFAULT_KEY_FIELD)
    aparser.add_argument('--show-type-fields', dest='show_type_fields', action='store_true', default=False)
    aparser.add_argument('--type-fields', dest='type_fields', action='store', default=','.join(DEFAULT_TYPE_FIELDS))
    aparser.add_argument('--all-type-fields', dest='all_type_fields', action='store_true', default=False)
    args = aparser.parse_args(argv[1:])
    filter_dict = reduce(lambda x,y:dict(x, **y),
                         (component_data_dict(file_name, 'Filter', 'Filters') for file_name in args.files),
                         {})
    fields = sorted(list(reduce(lambda x,y:x|y,
                                (set(v.keys()) for v in filter_dict.values()),
                                set())))
    if not args.all_fields:
        fields = [f.strip() for f in args.fields.split(',') if f.strip() in fields]
        if args.show_type_fields and 'Type' not in fields:
            fields.append('Type')
    #
    if 'Flags' in fields and not args.all_flags:
        for d in filter_dict.values():
            if 'Flags' in d:
                d['Flags'] = ' '.join([f for f in d['Flags'].split() if f in args.flags])
    #
    if args.show_type_fields:
        type_dict = reduce(lambda x,y:dict(x, **y),
                           (component_data_dict(file_name, 'Types', 'Types') for file_name in args.files),
                           {})
        type_fields = sorted(list(reduce(lambda x,y:x|y,
                                         (set(v.keys()) for v in type_dict.values()),
                                         set())))
        if not args.all_type_fields:
            type_fields = [f.strip() for f in args.type_fields.split(',') if f.strip() in type_fields]
    else:
        type_dict = {}
        type_fields = []
    #
    head = ['Name'] + fields + type_fields
    try:
        keyidx = head.index(args.key_field)
        key = lambda r:(r[keyidx], r)
    except:
        key = None
    table = sorted([[k] + [v.get(f) for f in fields] + [type_dict.get(v.get('Type'), {}).get(n) for n in type_fields] for k, v in filter_dict.items()],
                   key=key)
    table.insert(0, head)
    #
    if args.out is not None:
        outfp = open(args.out, 'w', newline='')
    else:
        outfp = sys.stdout
    csv.writer(outfp).writerows(table)
    if args.out is not None:
        outfp.close()
    return 0

if __name__ == '__main__':
    rc = main(sys.argv)
    sys.exit(rc)